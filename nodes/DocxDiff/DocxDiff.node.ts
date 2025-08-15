import type { IExecuteFunctions } from 'n8n-workflow';
import { NodeConnectionType, NodeOperationError } from 'n8n-workflow';
import type { INodeExecutionData, INodeType, INodeTypeDescription } from 'n8n-workflow';

import JSZip from 'jszip';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

function nowIso() {
	return new Date().toISOString();
}

type ResourceLimits = {
	maxTotalUnzippedBytes: number;
	maxEntries: number;
	maxEntrySize: number;
	maxTokensPerParagraph: number;
};

type Options = {
	granularity: 'word' | 'char';
	suppressWhitespaceOnly: boolean;
	includeLists: boolean;
	includeTables: boolean;
	includeTextBoxes: boolean;
	includeHeadersFooters: boolean;
	existingTrackedRevisions: 'ignore' | 'fail';
	limits: ResourceLimits;
};

type DiffInput = {
	base: Buffer;
	revised: Buffer;
	author: string;
	options: Options;
};

type DiffResult = {
	buffer: Buffer;
	warnings: string[];
};

const defaultLimits: ResourceLimits = {
	maxTotalUnzippedBytes: 50 * 1024 * 1024,
	maxEntries: 2000,
	maxEntrySize: 5 * 1024 * 1024,
	maxTokensPerParagraph: 4000,
};

const parser = new XMLParser({
	ignoreAttributes: false,
	attributeNamePrefix: '',
	allowBooleanAttributes: true,
	preserveOrder: false,
	numberParseOptions: { leadingZeros: false, hex: false },
	removeNSPrefix: true,
});

const builder = new XMLBuilder({
	ignoreAttributes: false,
	attributeNamePrefix: '',
	suppressEmptyNode: true,
	format: false,
});

function safeLoadDocx(buf: Buffer, limits: ResourceLimits, warnings: string[]): Promise<JSZip> {
	return JSZip.loadAsync(buf).then((zip) => {
		const files = Object.values(zip.files);
		if (files.length > limits.maxEntries) throw new Error('DOCX too many entries');
		let approxTotal = 0;
		for (const f of files) {
			approxTotal += ((f as any)._data?.compressedSize as number) || 0;
			if (approxTotal > limits.maxTotalUnzippedBytes * 2) throw new Error('DOCX likely exceeds unzip cap');
		}
		return zip;
	});
}

function stripTrackedChanges(node: any): any {
	if (!node || typeof node !== 'object') return node;
	if (node.ins || node.del || node.moveFrom || node.moveTo) {
		const content = (node.ins || node.del || node.moveFrom || node.moveTo) as any;
		if (Array.isArray(content.r)) return { r: content.r };
		return content;
	}
	for (const k of Object.keys(node)) {
		const v = (node as any)[k];
		if (Array.isArray(v)) (node as any)[k] = v.map(stripTrackedChanges);
		else if (typeof v === 'object') (node as any)[k] = stripTrackedChanges(v);
	}
	return node;
}

function tokenizeWords(text: string): string[] {
	return text.split(/(\s+|\b)/).filter((s) => s !== '');
}

function diffTokens(a: string[], b: string[]) {
	const m = a.length, n = b.length;
	const dp = Array.from({ length: m + 1 }, () => new Array<number>(n + 1).fill(0));
	for (let i = m - 1; i >= 0; i--) {
		for (let j = n - 1; j >= 0; j--) {
			dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
		}
	}
	const ops: Array<{ type: 'eq' | 'ins' | 'del'; value: string }>=[];
	let i = 0, j = 0;
	while (i < m && j < n) {
		if (a[i] === b[j]) { ops.push({ type: 'eq', value: a[i] }); i++; j++; }
		else if (dp[i + 1][j] >= dp[i][j + 1]) { ops.push({ type: 'del', value: a[i++] }); }
		else { ops.push({ type: 'ins', value: b[j++] }); }
	}
	while (i < m) ops.push({ type: 'del', value: a[i++] });
	while (j < n) ops.push({ type: 'ins', value: b[j++] });
	return ops;
}

function buildInsRuns(text: string, author: string, id: number) {
	return { ins: { id: String(id), author, date: nowIso(), r: [{ t: { '@_xml:space': 'preserve', '#text': text } }] } } as any;
}

function buildDelRuns(text: string, author: string, id: number) {
	return { del: { id: String(id), author, date: nowIso(), r: [{ delText: { '@_xml:space': 'preserve', '#text': text } }] } } as any;
}

function mergeRuns(ops: Array<{ type: 'eq' | 'ins' | 'del'; value: string }>, author: string) {
	const outRuns: any[] = [];
	let insBuf = '';
	let delBuf = '';
	let nextId = 1;
	const flush = () => {
		if (delBuf) { outRuns.push(buildDelRuns(delBuf, author, nextId++)); delBuf = ''; }
		if (insBuf) { outRuns.push(buildInsRuns(insBuf, author, nextId++)); insBuf = ''; }
	};
	for (const op of ops) {
		if (op.type === 'eq') { flush(); outRuns.push({ r: { t: { '@_xml:space': 'preserve', '#text': op.value } } }); }
		if (op.type === 'ins') insBuf += op.value;
		if (op.type === 'del') delBuf += op.value;
	}
	flush();
	return outRuns;
}

function paragraphText(p: any): string {
	const runs: any[] = Array.isArray(p.r) ? p.r : p.r ? [p.r] : [];
	let s = '';
	for (const r of runs) {
		if (typeof r.t === 'string') s += r.t;
		else if (r.t && typeof r.t['#text'] === 'string') s += r.t['#text'];
	}
	return s;
}

function diffParagraph(baseP: any, revP: any, author: string, options: Options): any {
	const a = paragraphText(baseP);
	const b = paragraphText(revP);
	if (options.suppressWhitespaceOnly && a.trim() === b.trim()) {
		return revP;
	}
	const tokensA = options.granularity === 'char' ? a.split('') : tokenizeWords(a);
	const tokensB = options.granularity === 'char' ? b.split('') : tokenizeWords(b);
	const ops = diffTokens(tokensA, tokensB);
	const runs = mergeRuns(ops, author);
	const pOut: any = { ...revP, r: undefined };
	pOut.r = runs.map((x) => x.r ?? x.ins ?? x.del);
	return pOut;
}

async function docxDiffTracked(input: DiffInput): Promise<DiffResult> {
	const warnings: string[] = [];
	const limits = input.options.limits ?? defaultLimits;
	let baseZip: JSZip; let revZip: JSZip;
	try {
		[baseZip, revZip] = await Promise.all([
			safeLoadDocx(input.base, limits, warnings),
			safeLoadDocx(input.revised, limits, warnings),
		]);
	} catch (e: any) {
		throw new Error(`Failed to read DOCX: ${e.message || e}`);
	}

	async function loadXml(z: JSZip, path: string): Promise<any | null> {
		const f = z.file(path);
		if (!f) return null;
		const xml = await f.async('string');
		try { return parser.parse(xml); } catch (e: any) { warnings.push(`Malformed ${path}; using raw fallback`); return null; }
	}

	const baseDoc = (await loadXml(baseZip, 'word/document.xml')) as any;
	const revDoc = (await loadXml(revZip, 'word/document.xml')) as any;
	if (!baseDoc || !revDoc) throw new Error('Missing word/document.xml in one of the DOCX files');

	if (input.options.existingTrackedRevisions === 'fail') {
		const hasTracked = JSON.stringify(revDoc).includes('ins') || JSON.stringify(revDoc).includes('del');
		if (hasTracked) throw new Error('Revised document contains tracked revisions');
	}

	const cleanBase = stripTrackedChanges(baseDoc);
	const cleanRev = stripTrackedChanges(revDoc);

	const baseParas: any[] = Array.isArray(cleanBase.document?.body?.p)
		? cleanBase.document.body.p
		: cleanBase.document?.body?.p ? [cleanBase.document.body.p] : [];
	const revParas: any[] = Array.isArray(cleanRev.document?.body?.p)
		? cleanRev.document.body.p
		: cleanRev.document?.body?.p ? [cleanRev.document.body.p] : [];

	const maxLen = Math.max(baseParas.length, revParas.length);
	const outParas: any[] = [];
	for (let i = 0; i < maxLen; i++) {
		const bp = baseParas[i];
		const rp = revParas[i];
		if (!bp && rp) {
			const text = paragraphText(rp);
			outParas.push({ p: { r: [buildInsRuns(text, input.author, i + 1).ins.r[0]] } });
			continue;
		}
		if (bp && !rp) {
			const text = paragraphText(bp);
			outParas.push({ p: { r: [buildDelRuns(text, input.author, i + 1).del.r[0]] } });
			continue;
		}
		if (bp && rp) {
			const d = diffParagraph(bp, rp, input.author, input.options);
			outParas.push({ p: d });
		}
	}

	const docOut: any = {
		document: {
			'@_xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
			body: {
				p: outParas.map((x) => x.p),
				sectPr: {},
			},
		},
	};
	const documentXml = builder.build(docOut);

	const settings: any = {
		settings: {
			'@_xmlns:w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
			trackRevisions: {},
		},
	};
	const settingsXml = builder.build(settings);

	const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`;

	const relsRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

	const zip = new JSZip();
	zip.file('[Content_Types].xml', contentTypes);
	zip.folder('_rels')?.file('.rels', relsRels);
	const w = zip.folder('word');
	w?.file('document.xml', documentXml);
	w?.file('settings.xml', settingsXml);

	const buffer = await zip.generateAsync({ type: 'nodebuffer' });
	return { buffer, warnings };
}

export class DocxDiff implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'DOCX Track-Diff',
		name: 'docxDiff',
		group: ['transform'],
		version: 1,
		description: 'Generate a DOCX with tracked changes from two DOCX binaries',
		defaults: { name: 'DOCX Track-Diff' },
		inputs: [NodeConnectionType.Main],
		outputs: [NodeConnectionType.Main],
		credentials: [],
		properties: [
			{
				displayName: 'Base Binary Property',
				name: 'baseProperty',
				type: 'string',
				default: 'base',
				description: 'Name of the binary property containing the base DOCX',
			},
			{
				displayName: 'Revised Binary Property',
				name: 'revisedProperty',
				type: 'string',
				default: 'revised',
				description: 'Name of the binary property containing the revised DOCX',
			},
			{
				displayName: 'Author',
				name: 'author',
				type: 'string',
				default: 'AutoDiff',
			},
			{
				displayName: 'Output File Name',
				name: 'outputFileName',
				type: 'string',
				default: 'diff.docx',
			},
			{
				displayName: 'Advanced',
				name: 'advanced',
				type: 'collection',
				placeholder: 'Advanced options',
				default: {},
				options: [
					{ displayName: 'Diff Granularity', name: 'granularity', type: 'options', options: [ { name: 'Character', value: 'char' }, { name: 'Word', value: 'word' } ], default: 'word' },
					{ displayName: 'Existing Tracked Revisions', name: 'existingTrackedRevisions', type: 'options', options: [ { name: 'Fail', value: 'fail' }, { name: 'Ignore', value: 'ignore' } ], default: 'ignore' },
					{ displayName: 'Include Headers/Footers', name: 'includeHeadersFooters', type: 'boolean', default: false },
					{ displayName: 'Include Lists', name: 'includeLists', type: 'boolean', default: true },
					{ displayName: 'Include Tables', name: 'includeTables', type: 'boolean', default: true },
					{ displayName: 'Include Text Boxes', name: 'includeTextBoxes', type: 'boolean', default: true },
					{ displayName: 'Resource Limits', name: 'limits', type: 'fixedCollection', default: {}, options: [
						{ displayName: 'Limits', name: 'limitsValues', values: [
							{ displayName: 'Max Entries', name: 'maxEntries', type: 'number', default: defaultLimits.maxEntries },
							{ displayName: 'Max Tokens Per Paragraph', name: 'maxTokensPerParagraph', type: 'number', default: defaultLimits.maxTokensPerParagraph },
							{ displayName: 'Per-entry Cap (Bytes)', name: 'maxEntrySize', type: 'number', default: defaultLimits.maxEntrySize },
							{ displayName: 'Total Unzip Cap (Bytes)', name: 'maxTotalUnzippedBytes', type: 'number', default: defaultLimits.maxTotalUnzippedBytes },
						] },
					] },
					{ displayName: 'Suppress Whitespace-only Suggestions', name: 'suppressWhitespaceOnly', type: 'boolean', default: true },
				],
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnItems: INodeExecutionData[] = [];
		for (let i = 0; i < items.length; i++) {
			try {
				const baseProp = this.getNodeParameter('baseProperty', i) as string;
				const revisedProp = this.getNodeParameter('revisedProperty', i) as string;
				const author = (this.getNodeParameter('author', i) as string) || 'AutoDiff';
				const outputFileName = (this.getNodeParameter('outputFileName', i) as string) || 'diff.docx';
				const adv = (this.getNodeParameter('advanced', i, {}) as any) || {};
				const limitsGroup = adv.limits?.limitsValues?.[0] ?? {};
				const options: Options = {
					granularity: (adv.granularity as any) ?? 'word',
					suppressWhitespaceOnly: adv.suppressWhitespaceOnly ?? true,
					includeLists: adv.includeLists ?? true,
					includeTables: adv.includeTables ?? true,
					includeTextBoxes: adv.includeTextBoxes ?? true,
					includeHeadersFooters: adv.includeHeadersFooters ?? false,
					existingTrackedRevisions: (adv.existingTrackedRevisions as any) ?? 'ignore',
					limits: {
						maxTotalUnzippedBytes: limitsGroup.maxTotalUnzippedBytes ?? defaultLimits.maxTotalUnzippedBytes,
						maxEntries: limitsGroup.maxEntries ?? defaultLimits.maxEntries,
						maxEntrySize: limitsGroup.maxEntrySize ?? defaultLimits.maxEntrySize,
						maxTokensPerParagraph: limitsGroup.maxTokensPerParagraph ?? defaultLimits.maxTokensPerParagraph,
					},
				};

				const baseBuf = await this.helpers.getBinaryDataBuffer(i, baseProp);
				const revBuf = await this.helpers.getBinaryDataBuffer(i, revisedProp);
				if (!baseBuf) throw new NodeOperationError(this.getNode(), `Binary property "${baseProp}" not found`, { itemIndex: i });
				if (!revBuf) throw new NodeOperationError(this.getNode(), `Binary property "${revisedProp}" not found`, { itemIndex: i });

				const { buffer, warnings } = await docxDiffTracked({ base: baseBuf, revised: revBuf, author, options });
				const binary = await this.helpers.prepareBinaryData(
					buffer,
					outputFileName,
					'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				);
				returnItems.push({ json: { warnings }, binary: binary as any });
			} catch (error: any) {
				if (this.continueOnFail()) {
					returnItems.push({ json: { error: error.message || String(error), warnings: [] } });
					continue;
				}
				throw new NodeOperationError(this.getNode(), error, { itemIndex: i });
			}
		}
		return [returnItems];
	}
}
