import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import mammoth from 'mammoth';
import {
    AlignCenter, AlignLeft, AlignRight, Bold, Check, ChevronDown, Clipboard, Code2,
    Download, FileCode2, FileText, Image as ImageIcon,
    Italic, Link2, List, ListOrdered, LoaderCircle, Menu, Minus, Paintbrush,
    Pilcrow, Quote, RotateCcw, Settings2, Sparkles, Strikethrough,
    Underline, Upload, WandSparkles, X, FileUp, Eye, Copy, CheckCircle2,
} from 'lucide-react';
import {
    cleanHtml, copyText, countWords, downloadFile, embedImagesAsBase64,
    formatHtml, htmlToMarkdown, htmlToText, wrapHtmlDocument,
} from './utils.js';

const DEFAULT_HTML = `<h2>Welcome to Word to HTML Studio</h2>
<p>Upload a <strong>Word document (.docx)</strong> or paste content copied from Word to get started.</p>
<ul>
  <li>Switch between a visual editor and editable HTML source.</li>
  <li>Keep document images embedded as base64 data URLs.</li>
  <li>Apply cleanup rules and export clean HTML, text, Markdown, or JSON.</li>
</ul>
<p>Try editing this sample content, or import your own document.</p>`;

const INITIAL_CLEANUP = {
    removeInlineStyles: false,
    removeClassesIds: true,
    removeEmptyTags: true,
    removeBold: false,
    removeItalic: false,
    removeUnderline: false,
    removeLinks: false,
    removeFontTags: false,
    removeComments: true,
    stripFormatting: false,
};

const CLEANUP_OPTIONS = [
    { key: 'removeInlineStyles', label: 'Remove inline styles', detail: 'Strip style attributes from all elements.' },
    { key: 'removeClassesIds', label: 'Remove classes and IDs', detail: 'Remove class and id attributes.' },
    { key: 'removeEmptyTags', label: 'Remove empty tags', detail: 'Delete empty paragraphs and empty elements.' },
    { key: 'removeBold', label: 'Remove bold formatting', detail: 'Keep text, unwrap <strong> and <b>.' },
    { key: 'removeItalic', label: 'Remove italic formatting', detail: 'Keep text, unwrap <em> and <i>.' },
    { key: 'removeUnderline', label: 'Remove underline formatting', detail: 'Keep text, unwrap <u>.' },
    { key: 'removeLinks', label: 'Remove hyperlinks', detail: 'Keep link text and remove the link target.' },
    { key: 'removeFontTags', label: 'Remove font and span tags', detail: 'Unwrap font, span, and minor formatting tags.' },
    { key: 'removeComments', label: 'Remove HTML comments', detail: 'Delete comments from the source.' },
    { key: 'stripFormatting', label: 'Strip formatting', detail: 'Remove common inline formatting and convert headings to paragraphs.' },
];

function Toggle({ checked, onChange, label }) {
    return (
        <button
            type="button"
            role="switch"
            aria-checked={checked}
            aria-label={label}
            onClick={() => onChange(!checked)}
            className={`toggle ${checked ? 'toggle-on' : ''}`}
        >
            <span className="toggle-knob" />
        </button>
    );
}

function ToolButton({ title, onClick, children, active = false, disabled = false }) {
    return (
        <button
            type="button"
            title={title}
            aria-label={title}
            disabled={disabled}
            className={`editor-tool ${active ? 'editor-tool-active' : ''}`}
            onMouseDown={(event) => event.preventDefault()}
            onClick={onClick}
        >
            {children}
        </button>
    );
}

function normalizeLinkUrl(value) {
    const url = value.trim();
    if (!url) return '';
    // Preserve mail, telephone, anchor, query, root-relative, and protocol-relative URLs.
    if (/^(?:[a-z][a-z\d+.-]*:|\/|#|\?)/i.test(url)) return url;
    return `https://${url}`;
}

function PasteModal({ onClose, onConvert }) {
    const pasteRef = useRef(null);
    const [pasteMode, setPasteMode] = useState('visual');
    const [rawText, setRawText] = useState('');
    const [hasContent, setHasContent] = useState(false);

    const convert = () => {
        const html = pasteMode === 'visual'
            ? (pasteRef.current?.innerHTML || '')
            : rawText.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').split(/\n\s*\n/).map((p) => `<p>${p.replace(/\n/g, '<br>')}</p>`).join('\n');
        onConvert(html);
    };

    const onPaste = async (event) => {
        const items = [...(event.clipboardData?.items || [])];
        const imageFiles = items.filter((item) => item.type.startsWith('image/')).map((item) => item.getAsFile()).filter(Boolean);
        // Browsers usually supply Word's rich HTML directly. For image-only clipboard data, create data URLs ourselves.
        if (imageFiles.length && !event.clipboardData?.getData('text/html')) {
            event.preventDefault();
            for (const file of imageFiles) {
                const dataUrl = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result);
                    reader.onerror = reject;
                    reader.readAsDataURL(file);
                });
                pasteRef.current?.insertAdjacentHTML('beforeend', `<img src="${dataUrl}" alt="Pasted image">`);
            }
        }
        window.setTimeout(() => setHasContent(Boolean(pasteRef.current?.innerHTML?.trim())), 0);
    };

    return (
        <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
            <section className="paste-modal" role="dialog" aria-modal="true" aria-labelledby="paste-title">
                <div className="modal-header">
                    <div className="modal-icon"><Clipboard size={19} /></div>
                    <div>
                        <h2 id="paste-title">Paste content from Word</h2>
                        <p>Copy text, tables, formatting, and images from Word, then paste below.</p>
                    </div>
                    <button className="icon-button modal-close" onClick={onClose} aria-label="Close paste dialog"><X size={19} /></button>
                </div>
                <div className="paste-tabs">
                    <button className={pasteMode === 'visual' ? 'paste-tab active' : 'paste-tab'} onClick={() => setPasteMode('visual')}>Rich content</button>
                    <button className={pasteMode === 'text' ? 'paste-tab active' : 'paste-tab'} onClick={() => setPasteMode('text')}>Plain text</button>
                </div>
                {pasteMode === 'visual' ? (
                    <div
                        ref={pasteRef}
                        className="paste-canvas"
                        contentEditable
                        suppressContentEditableWarning
                        onInput={() => setHasContent(Boolean(pasteRef.current?.innerHTML?.trim()))}
                        onPaste={onPaste}
                        data-placeholder="Click here, then press Ctrl + V to paste from Word…"
                    />
                ) : (
                    <textarea
                        className="paste-textarea"
                        placeholder="Paste plain text here…"
                        value={rawText}
                        onChange={(event) => setRawText(event.target.value)}
                    />
                )}
                <div className="paste-note"><ImageIcon size={15} /> Images copied with Word's clipboard may depend on browser permissions. DOCX upload is best for reliably embedding every document image.</div>
                <div className="modal-footer">
                    <button className="button button-quiet" onClick={() => { if (pasteRef.current) pasteRef.current.innerHTML = ''; setRawText(''); setHasContent(false); }}>Clear</button>
                    <div className="modal-footer-right">
                        <button className="button button-quiet" onClick={onClose}>Cancel</button>
                        <button className="button button-primary" onClick={convert} disabled={pasteMode === 'visual' ? !hasContent : !rawText.trim()}><WandSparkles size={16} /> Convert pasted content</button>
                    </div>
                </div>
            </section>
        </div>
    );
}

function App() {
    const [sourceHtml, setSourceHtml] = useState(DEFAULT_HTML);
    const [activeView, setActiveView] = useState('visual');
    const [codeText, setCodeText] = useState(formatHtml(DEFAULT_HTML));
    const [codeDirty, setCodeDirty] = useState(false);
    const [cleanup, setCleanup] = useState(INITIAL_CLEANUP);
    const [showCleanup, setShowCleanup] = useState(true);
    const [showPaste, setShowPaste] = useState(false);
    const [busy, setBusy] = useState(false);
    const [fileName, setFileName] = useState('Untitled document');
    const [notice, setNotice] = useState({ type: 'info', text: 'Ready to convert. Images in uploaded DOCX files are embedded as base64.' });
    const [exportMenuPosition, setExportMenuPosition] = useState(null);
    const [showMobileMenu, setShowMobileMenu] = useState(false);
    const [copied, setCopied] = useState(false);
    const [activeBlock, setActiveBlock] = useState('p');
    const [activeFormatting, setActiveFormatting] = useState({
        bold: false,
        italic: false,
        underline: false,
        strikeThrough: false,
        unorderedList: false,
        orderedList: false,
    });
    const [linkEditor, setLinkEditor] = useState({ open: false, url: '', top: 0, left: 8 });
    const editorRef = useRef(null);
    const fileInputRef = useRef(null);
    const dropRef = useRef(null);
    const linkInputRef = useRef(null);
    const linkAnchorRef = useRef(null);
    const linkRangeRef = useRef(null);

    const processedHtml = useMemo(() => cleanHtml(sourceHtml, cleanup), [sourceHtml, cleanup]);
    const outputText = useMemo(() => htmlToText(processedHtml), [processedHtml]);
    const wordCount = useMemo(() => countWords(processedHtml), [processedHtml]);
    const imageCount = useMemo(() => (processedHtml.match(/<img\b/gi) || []).length, [processedHtml]);
    const paragraphCount = useMemo(() => (processedHtml.match(/<(?:p|h[1-6])\b/gi) || []).length, [processedHtml]);
    const chars = outputText.length;

    useEffect(() => {
        if (activeView === 'visual' && editorRef.current) {
            const current = editorRef.current.innerHTML;
            if (current !== processedHtml) editorRef.current.innerHTML = processedHtml;
        }
    }, [processedHtml, activeView]);

    useEffect(() => {
        if (activeView === 'code' && !codeDirty) setCodeText(formatHtml(processedHtml));
    }, [processedHtml, activeView, codeDirty]);

    useEffect(() => {
        if (linkEditor.open) linkInputRef.current?.focus();
    }, [linkEditor.open]);

    useEffect(() => {
        if (activeView !== 'visual') return undefined;
        const handleSelectionChange = () => syncToolbarState();
        document.addEventListener('selectionchange', handleSelectionChange);
        return () => document.removeEventListener('selectionchange', handleSelectionChange);
    }, [activeView]);

    const setNoticeSafe = useCallback((type, text) => setNotice({ type, text }), []);

    const syncToolbarState = (sourceNode = null) => {
        const selection = window.getSelection();
        const node = sourceNode || selection?.anchorNode;
        const element = node?.nodeType === Node.ELEMENT_NODE ? node : node?.parentElement;
        if (!element || !editorRef.current?.contains(element)) return;

        const list = element.closest('ul, ol');
        const block = list || element.closest('h1, h2, h3, h4, h5, h6, p, blockquote, pre');
        const blockTag = block?.tagName?.toLowerCase();
        setActiveBlock(blockTag && ['ul', 'ol', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'blockquote', 'pre'].includes(blockTag) ? blockTag : 'p');

        const inTag = (selector) => Boolean(element.closest(selector));
        const commandState = (command) => {
            try { return document.queryCommandState(command); } catch { return false; }
        };
        setActiveFormatting({
            bold: commandState('bold') || inTag('strong, b'),
            italic: commandState('italic') || inTag('em, i'),
            underline: commandState('underline') || inTag('u'),
            strikeThrough: commandState('strikeThrough') || inTag('s, strike, del'),
            unorderedList: list?.tagName?.toLowerCase() === 'ul',
            orderedList: list?.tagName?.toLowerCase() === 'ol',
        });
    };

    const processDocx = useCallback(async (file) => {
        if (!file) return;
        if (!/\.docx$/i.test(file.name) && file.type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            setNoticeSafe('error', 'Please choose a .docx file. The older .doc format is not supported by in-browser conversion; save it as .docx in Word first.');
            return;
        }
        setBusy(true);
        setNoticeSafe('info', `Reading ${file.name}…`);
        try {
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.convertToHtml({ arrayBuffer }, {
                convertImage: mammoth.images.imgElement((image) => image.read('base64').then((imageBuffer) => ({
                    src: `data:${image.contentType};base64,${imageBuffer}`,
                }))),
            });
            const docImageCount = ((result.value || '').match(/<img\b/gi) || []).length;
            setSourceHtml(result.value || '<p></p>');
            setFileName(file.name.replace(/\.docx$/i, '') || 'Converted document');
            setActiveView('visual');
            setCodeDirty(false);
            const warningCount = result.messages.filter((message) => message.type === 'warning').length;
            setNoticeSafe(warningCount ? 'warning' : 'success', `Converted ${file.name}. ${docImageCount} image${docImageCount === 1 ? '' : 's'} embedded as base64.${warningCount ? ` Mammoth reported ${warningCount} conversion note(s); review complex Word layouts in the editor.` : ''}`);
        } catch (error) {
            console.error(error);
            setNoticeSafe('error', 'This DOCX could not be converted. Check that it is a valid, unencrypted Word document and try again.');
        } finally {
            setBusy(false);
            if (fileInputRef.current) fileInputRef.current.value = '';
        }
    }, [setNoticeSafe]);

    const importPastedHtml = useCallback(async (html) => {
        setBusy(true);
        setShowPaste(false);
        setNoticeSafe('info', 'Cleaning pasted content and embedding reachable images…');
        try {
            const result = await embedImagesAsBase64(html);
            setSourceHtml(result.html || '<p></p>');
            setFileName('Pasted Word content');
            setActiveView('visual');
            setCodeDirty(false);
            const message = `Imported pasted content${result.embeddedCount ? `; embedded ${result.embeddedCount} image${result.embeddedCount === 1 ? '' : 's'}` : ''}.`;
            setNoticeSafe(result.failedCount ? 'warning' : 'success', `${message}${result.failedCount ? ` ${result.failedCount} image(s) could not be embedded, often because the source is local or blocks cross-origin access.` : ''}`);
        } catch (error) {
            console.error(error);
            setNoticeSafe('error', 'Pasted content could not be processed. Try pasting again or uploading the original DOCX.');
        } finally {
            setBusy(false);
        }
    }, [setNoticeSafe]);

    const changeView = (view) => {
        if (view === activeView) return;
        if (view === 'code') {
            setCodeText(formatHtml(processedHtml));
            setCodeDirty(false);
            setActiveView('code');
            return;
        }
        if (activeView === 'code') {
            setSourceHtml(codeText);
            setCodeDirty(false);
        } else if (editorRef.current) {
            setSourceHtml(editorRef.current.innerHTML);
        }
        setActiveView(view);
    };

    const applyCode = () => {
        setSourceHtml(codeText);
        setCodeDirty(false);
        setNoticeSafe('success', 'HTML source applied to the live editor.');
    };

    const handleEditorInput = () => {
        if (editorRef.current) setSourceHtml(editorRef.current.innerHTML);
        syncToolbarState();
    };

    const runCommand = (command, value = null) => {
        if (!editorRef.current) return;
        editorRef.current.focus();
        document.execCommand(command, false, value);
        setSourceHtml(editorRef.current.innerHTML);
        syncToolbarState();
    };

    const changeBlock = (tag) => {
        if (!tag) return;
        if (tag === 'ul') {
            runCommand('insertUnorderedList');
            return;
        }
        if (tag === 'ol') {
            runCommand('insertOrderedList');
            return;
        }
        runCommand('formatBlock', `<${tag}>`);
    };

    const openLinkEditor = (anchor = null, providedRange = null) => {
        const editor = editorRef.current;
        const wrapper = editor?.closest('.visual-editor-wrap');
        if (!editor || !wrapper) return;

        const selection = window.getSelection();
        const range = providedRange || (selection?.rangeCount ? selection.getRangeAt(0).cloneRange() : null);
        linkAnchorRef.current = anchor;
        linkRangeRef.current = range;

        let rect = anchor?.getBoundingClientRect();
        if (!rect && range) rect = range.getBoundingClientRect();
        if ((!rect || (!rect.width && !rect.height)) && range?.startContainer?.parentElement) {
            rect = range.startContainer.parentElement.getBoundingClientRect();
        }
        if (!rect) rect = editor.getBoundingClientRect();

        const wrapperRect = wrapper.getBoundingClientRect();
        const popoverWidth = 320;
        const left = Math.min(Math.max(8, rect.left - wrapperRect.left), Math.max(8, wrapperRect.width - popoverWidth - 8));
        const top = Math.max(8, rect.bottom - wrapperRect.top + 6);
        setLinkEditor({ open: true, url: anchor?.getAttribute('href') || '', top, left });
    };

    const insertLink = () => {
        const selection = window.getSelection();
        const node = selection?.anchorNode;
        const element = node?.nodeType === Node.ELEMENT_NODE ? node : node?.parentElement;
        const anchor = element && editorRef.current?.contains(element) ? element.closest('a') : null;
        const range = selection?.rangeCount && editorRef.current?.contains(element) ? selection.getRangeAt(0).cloneRange() : null;
        openLinkEditor(anchor, range);
    };

    const closeLinkEditor = () => {
        setLinkEditor((current) => ({ ...current, open: false }));
        linkAnchorRef.current = null;
        linkRangeRef.current = null;
    };

    const saveLink = () => {
        const editor = editorRef.current;
        const anchor = linkAnchorRef.current;
        const url = normalizeLinkUrl(linkEditor.url);
        if (!editor) return closeLinkEditor();

        if (anchor && editor.contains(anchor)) {
            if (!url) {
                const parent = anchor.parentNode;
                if (parent) {
                    while (anchor.firstChild) parent.insertBefore(anchor.firstChild, anchor);
                    parent.removeChild(anchor);
                }
                setNoticeSafe('success', 'Hyperlink removed; the text was kept.');
            } else {
                anchor.setAttribute('href', url);
                setNoticeSafe('success', 'Hyperlink updated.');
            }
        } else if (url) {
            editor.focus();
            const selection = window.getSelection();
            const range = linkRangeRef.current;
            if (range && selection) {
                selection.removeAllRanges();
                selection.addRange(range);
                if (!range.collapsed) {
                    document.execCommand('createLink', false, url);
                } else {
                    const newAnchor = document.createElement('a');
                    newAnchor.href = url;
                    newAnchor.textContent = linkEditor.url.trim();
                    range.insertNode(newAnchor);
                    range.setStartAfter(newAnchor);
                    range.collapse(true);
                    selection.removeAllRanges();
                    selection.addRange(range);
                }
                setNoticeSafe('success', 'Hyperlink inserted.');
            } else {
                setNoticeSafe('warning', 'Select some text or place the cursor in the editor before adding a link.');
            }
        }

        setSourceHtml(editor.innerHTML);
        closeLinkEditor();
        window.requestAnimationFrame(() => syncToolbarState());
    };

    const handleEditorClick = (event) => {
        const anchor = event.target.closest?.('a');
        if (anchor && editorRef.current?.contains(anchor)) {
            event.preventDefault();
            openLinkEditor(anchor);
            syncToolbarState(event.target);
            return;
        }
        syncToolbarState(event.target);
    };

    const copyHtml = async () => {
        try {
            await copyText(formatHtml(processedHtml));
            setCopied(true);
            setNoticeSafe('success', 'HTML copied to the clipboard.');
            window.setTimeout(() => setCopied(false), 1800);
        } catch {
            setNoticeSafe('error', 'Clipboard access was blocked. Use the HTML view to select and copy the source.');
        }
    };

    const exportAs = (type) => {
        const base = (fileName || 'converted-document').replace(/[^a-z0-9-_]+/gi, '-').replace(/^-|-$/g, '') || 'converted-document';
        const fragment = formatHtml(processedHtml);
        if (type === 'html') downloadFile(`${base}.html`, wrapHtmlDocument(processedHtml, fileName), 'text/html;charset=utf-8');
        if (type === 'fragment') downloadFile(`${base}-fragment.html`, fragment, 'text/html;charset=utf-8');
        if (type === 'txt') downloadFile(`${base}.txt`, outputText, 'text/plain;charset=utf-8');
        if (type === 'md') downloadFile(`${base}.md`, htmlToMarkdown(processedHtml), 'text/markdown;charset=utf-8');
        if (type === 'json') downloadFile(`${base}.json`, JSON.stringify({ fileName, html: processedHtml, text: outputText, markdown: htmlToMarkdown(processedHtml), images: imageCount, words: wordCount }, null, 2), 'application/json;charset=utf-8');
        setExportMenuPosition(null);
        setNoticeSafe('success', `Exported ${type === 'fragment' ? 'HTML fragment' : type.toUpperCase()} file.`);
    };

    const resetDocument = () => {
        if (sourceHtml !== DEFAULT_HTML && !window.confirm('Replace the current document with the starter sample?')) return;
        setSourceHtml(DEFAULT_HTML);
        setFileName('Untitled document');
        setActiveView('visual');
        setCodeText(formatHtml(DEFAULT_HTML));
        setCodeDirty(false);
        setCleanup(INITIAL_CLEANUP);
        setNoticeSafe('info', 'Starter sample restored.');
    };

    const handleDrop = async (event) => {
        event.preventDefault();
        dropRef.current?.classList.remove('drag-over');
        const file = event.dataTransfer.files?.[0];
        if (file) await processDocx(file);
    };

    const handleDragOver = (event) => {
        event.preventDefault();
        dropRef.current?.classList.add('drag-over');
    };

    const handleDragLeave = () => dropRef.current?.classList.remove('drag-over');

    return (
        <div className="app-shell">
            <header className="topbar">
                <div className="brand-lockup">
                    <div className="brand-mark">
                        {/* <FileCode2 size={21} strokeWidth={2.2} /> */}
                        <img src="/logo.png" alt="Word to HTML Studio" className="w-32 mt-4" />
                    </div>
                    {/* <div className="brand-text"><div className="brand-name">Word to HTML <span>Studio</span></div><div className="brand-subtitle">Convert · Clean · Export</div></div> */}
                </div>
                <div className="topbar-actions">
                    <div className="local-badge"><span className="local-dot" /> Runs in your browser</div>
                    <button className="button button-quiet reset-button" onClick={resetDocument}><RotateCcw size={15} /> Reset</button>
                    <button className="button button-primary top-export" onClick={() => setExportMenuPosition((value) => value === 'top' ? null : 'top')}><Download size={16} /> Export <ChevronDown size={15} /></button>
                    {exportMenuPosition === 'top' && <div className="export-menu top-export-menu">
                        <button onClick={() => exportAs('html')}><FileCode2 size={16} /><span>HTML document</span><small>.html</small></button>
                        <button onClick={() => exportAs('fragment')}><Code2 size={16} /><span>HTML fragment</span><small>.html</small></button>
                        <button onClick={() => exportAs('txt')}><FileText size={16} /><span>Plain text</span><small>.txt</small></button>
                        <button onClick={() => exportAs('md')}><Pilcrow size={16} /><span>Markdown</span><small>.md</small></button>
                        <button onClick={() => exportAs('json')}><FileText size={16} /><span>HTML + text bundle</span><small>.json</small></button>
                    </div>}
                    <button className="icon-button mobile-menu-trigger" aria-label="Toggle mobile controls" onClick={() => setShowMobileMenu((value) => !value)}><Menu size={20} /></button>
                </div>
            </header>

            <main className="workspace">
                <section className="intro-row">
                    <div>
                        <div className="eyebrow"><Sparkles size={14} /> DOCUMENT CONVERSION WORKSPACE</div>
                        <h1>Turn Word documents into <span>clean HTML.</span></h1>
                        <p className="intro-description">Convert DOCX files, edit the result visually or in source, clean up markup, and export in the format you need.</p>
                    </div>
                    <div className="intro-shortcut"><span className="shortcut-icon"><ImageIcon size={18} /></span><div><strong>Base64 image embedding</strong><span>DOCX images travel with your HTML</span></div></div>
                </section>

                <div className={`mobile-actions ${showMobileMenu ? 'mobile-actions-open' : ''}`}>
                    <button className="button button-primary" onClick={() => fileInputRef.current?.click()}><Upload size={16} /> Upload DOCX</button>
                    <button className="button button-secondary" onClick={() => setShowPaste(true)}><Clipboard size={16} /> Paste from Word</button>
                </div>

                <section
                    className="import-strip"
                    ref={dropRef}
                    onDrop={handleDrop}
                    onDragOver={handleDragOver}
                    onDragLeave={handleDragLeave}
                >
                    <input ref={fileInputRef} className="hidden-input" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onChange={(event) => processDocx(event.target.files?.[0])} />
                    <div className="import-main">
                        <div className="upload-icon"><FileUp size={22} /></div>
                        <div className="import-copy"><strong>Start with a Word document</strong><span>Drag a .docx file here or choose one from your computer.</span></div>
                    </div>
                    <div className="import-actions">
                        <button className="button button-secondary" onClick={() => setShowPaste(true)}><Clipboard size={16} /> Paste from Word</button>
                        <button className="button button-primary" onClick={() => fileInputRef.current?.click()} disabled={busy}>{busy ? <LoaderCircle size={16} className="spin" /> : <Upload size={16} />} Upload DOCX</button>
                    </div>
                </section>

                <section className="notice-row" aria-live="polite">
                    <span className={`notice-icon notice-${notice.type}`}>{notice.type === 'success' ? <CheckCircle2 size={16} /> : notice.type === 'error' ? <X size={16} /> : notice.type === 'warning' ? <Paintbrush size={16} /> : <Sparkles size={15} />}</span>
                    <span>{notice.text}</span>
                    {busy && <LoaderCircle className="notice-spinner spin" size={15} />}
                </section>

                <div className={`main-layout ${!showCleanup ? 'main-layout-full' : ''}`}>
                    <section className="editor-panel">
                        <div className="panel-header">
                            <div className="doc-title-wrap"><div className="doc-title-icon"><FileText size={17} /></div><div className="doc-title"><strong>{fileName}</strong><span>{activeView === 'visual' ? 'Visual editor' : 'HTML source'} · unsaved local workspace</span></div></div>
                            <div className="panel-header-actions">
                                <button className="small-icon-button" onClick={copyHtml} title="Copy HTML" aria-label="Copy HTML"><Copy size={16} /></button>
                                <button className="small-icon-button" onClick={() => setShowCleanup((value) => !value)} title="Toggle cleanup panel" aria-label="Toggle cleanup panel"><Settings2 size={17} /></button>
                            </div>
                        </div>

                        <div className="view-tabs" role="tablist" aria-label="Editor view">
                            <button role="tab" aria-selected={activeView === 'visual'} className={`view-tab ${activeView === 'visual' ? 'view-tab-active' : ''}`} onClick={() => changeView('visual')}><Eye size={16} /> Live view <span className="tab-tag">EDIT</span></button>
                            <button role="tab" aria-selected={activeView === 'code'} className={`view-tab ${activeView === 'code' ? 'view-tab-active' : ''}`} onClick={() => changeView('code')}><Code2 size={16} /> HTML code</button>
                            <div className="view-tab-spacer" />
                            <button className={`copy-source-button ${copied ? 'copied' : ''}`} onClick={copyHtml}>{copied ? <Check size={14} /> : <Copy size={14} />} {copied ? 'Copied' : 'Copy HTML'}</button>
                        </div>

                        {activeView === 'visual' ? (
                            <div className="visual-editor-wrap">
                                <div className="classic-toolbar" role="toolbar" aria-label="Visual editor formatting">
                                    <select className="format-select" value={activeBlock} aria-label="Block format" onChange={(event) => changeBlock(event.target.value)}>
                                        <option value="p">Paragraph</option><option value="h1">Heading 1</option><option value="h2">Heading 2</option><option value="h3">Heading 3</option><option value="h4">Heading 4</option><option value="h5">Heading 5</option><option value="h6">Heading 6</option><option value="ul">Bulleted list</option><option value="ol">Numbered list</option><option value="blockquote">Quote</option><option value="pre">Preformatted</option>
                                    </select>
                                    <span className="toolbar-divider" />
                                    <ToolButton title="Bold" active={activeFormatting.bold} onClick={() => runCommand('bold')}><Bold size={16} /></ToolButton>
                                    <ToolButton title="Italic" active={activeFormatting.italic} onClick={() => runCommand('italic')}><Italic size={16} /></ToolButton>
                                    <ToolButton title="Underline" active={activeFormatting.underline} onClick={() => runCommand('underline')}><Underline size={16} /></ToolButton>
                                    <ToolButton title="Strikethrough" active={activeFormatting.strikeThrough} onClick={() => runCommand('strikeThrough')}><Strikethrough size={16} /></ToolButton>
                                    <span className="toolbar-divider" />
                                    <ToolButton title="Bulleted list" active={activeFormatting.unorderedList} onClick={() => runCommand('insertUnorderedList')}><List size={17} /></ToolButton>
                                    <ToolButton title="Numbered list" active={activeFormatting.orderedList} onClick={() => runCommand('insertOrderedList')}><ListOrdered size={17} /></ToolButton>
                                    <ToolButton title="Block quote" active={activeBlock === 'blockquote'} onClick={() => changeBlock('blockquote')}><Quote size={16} /></ToolButton>
                                    <ToolButton title="Insert link" onClick={insertLink}><Link2 size={16} /></ToolButton>
                                    <span className="toolbar-divider" />
                                    <ToolButton title="Align left" onClick={() => runCommand('justifyLeft')}><AlignLeft size={16} /></ToolButton>
                                    <ToolButton title="Align center" onClick={() => runCommand('justifyCenter')}><AlignCenter size={16} /></ToolButton>
                                    <ToolButton title="Align right" onClick={() => runCommand('justifyRight')}><AlignRight size={16} /></ToolButton>
                                    <ToolButton title="Horizontal line" onClick={() => runCommand('insertHorizontalRule')}><Minus size={16} /></ToolButton>
                                </div>
                                <div
                                    ref={editorRef}
                                    className="live-editor content-prose"
                                    contentEditable
                                    suppressContentEditableWarning
                                    spellCheck
                                    onInput={handleEditorInput}
                                    onClick={handleEditorClick}
                                    onKeyUp={() => syncToolbarState()}
                                    onMouseUp={() => syncToolbarState()}
                                    data-placeholder="Your converted document will appear here…"
                                />
                                {linkEditor.open && <div
                                    className="link-editor-popover"
                                    style={{ top: `${linkEditor.top}px`, left: `${linkEditor.left}px` }}
                                    role="group"
                                    aria-label="Edit hyperlink"
                                    onMouseDown={(event) => event.stopPropagation()}
                                >
                                    <label htmlFor="link-url-input">Link URL</label>
                                    <input
                                        ref={linkInputRef}
                                        id="link-url-input"
                                        type="text"
                                        value={linkEditor.url}
                                        placeholder="https://example.com"
                                        onChange={(event) => setLinkEditor((current) => ({ ...current, url: event.target.value }))}
                                        onKeyDown={(event) => {
                                            if (event.key === 'Enter') { event.preventDefault(); saveLink(); }
                                            if (event.key === 'Escape') { event.preventDefault(); closeLinkEditor(); }
                                        }}
                                    />
                                    <button className="button button-primary link-save-button" onClick={saveLink}><Check size={14} /> Save</button>
                                    <button className="icon-button link-close-button" aria-label="Close link editor" title="Close" onClick={closeLinkEditor}><X size={15} /></button>
                                </div>}
                                <div className="editor-bottomline"><span><span className="ready-dot" /> Visual editor</span><span>Click any text to edit · formatting toolbar above</span></div>
                            </div>
                        ) : (
                            <div className="code-editor-wrap">
                                <div className="code-notice"><Code2 size={15} /><span>Edit the HTML directly, then switch to Live view or select Apply changes.</span><span className="code-language">HTML</span></div>
                                <textarea
                                    className="code-editor"
                                    aria-label="Editable HTML source"
                                    spellCheck={false}
                                    value={codeText}
                                    onChange={(event) => { setCodeText(event.target.value); setCodeDirty(true); }}
                                />
                                <div className="code-actions"><span>{codeText.length.toLocaleString()} characters {codeDirty && <em>· unapplied changes</em>}</span><div><button className="button button-quiet" onClick={() => { setCodeText(formatHtml(processedHtml)); setCodeDirty(false); }}>Discard edits</button><button className="button button-primary" onClick={applyCode}><Check size={15} /> Apply changes</button></div></div>
                            </div>
                        )}

                        <div className="stats-bar">
                            <div><span className="stat-number">{wordCount.toLocaleString()}</span><span className="stat-label">Words</span></div>
                            <div><span className="stat-number">{paragraphCount.toLocaleString()}</span><span className="stat-label">Paragraphs & headings</span></div>
                            <div><span className="stat-number">{imageCount.toLocaleString()}</span><span className="stat-label">Images</span></div>
                            <div><span className="stat-number">{chars.toLocaleString()}</span><span className="stat-label">Text characters</span></div>
                            <div className="stats-spacer" />
                            <button className="stats-export" onClick={() => setExportMenuPosition((value) => value === 'bottom' ? null : 'bottom')}><Download size={15} /> Export <ChevronDown size={14} /></button>
                            {exportMenuPosition === 'bottom' && <div className="export-menu stats-export-menu">
                                <button onClick={() => exportAs('html')}><FileCode2 size={16} /><span>HTML document</span><small>.html</small></button>
                                <button onClick={() => exportAs('fragment')}><Code2 size={16} /><span>HTML fragment</span><small>.html</small></button>
                                <button onClick={() => exportAs('txt')}><FileText size={16} /><span>Plain text</span><small>.txt</small></button>
                                <button onClick={() => exportAs('md')}><Pilcrow size={16} /><span>Markdown</span><small>.md</small></button>
                                <button onClick={() => exportAs('json')}><FileText size={16} /><span>HTML + text bundle</span><small>.json</small></button>
                            </div>}
                        </div>
                    </section>

                    {showCleanup && <aside className="cleanup-panel">
                        <div className="cleanup-heading"><div className="cleanup-heading-icon"><WandSparkles size={17} /></div><div><h2>HTML cleanup</h2><p>Changes update the output live.</p></div><button className="icon-button cleanup-close" aria-label="Close cleanup panel" onClick={() => setShowCleanup(false)}><X size={16} /></button></div>
                        <div className="cleanup-callout"><Sparkles size={16} /><div><strong>Clean markup, your way</strong><span>Toggle rules on or off. The editor preview updates instantly.</span></div></div>
                        <div className="cleanup-options">
                            {CLEANUP_OPTIONS.map((option) => (
                                <div className="cleanup-option" key={option.key}>
                                    <div className="cleanup-option-copy"><label>{option.label}</label><p>{option.detail}</p></div>
                                    <Toggle checked={cleanup[option.key]} label={option.label} onChange={(checked) => setCleanup((current) => ({ ...current, [option.key]: checked }))} />
                                </div>
                            ))}
                        </div>
                        <div className="cleanup-footer"><button className="button button-quiet reset-cleanup" onClick={() => setCleanup(INITIAL_CLEANUP)}><RotateCcw size={14} /> Reset rules</button><span><span className="live-dot" /> Real-time</span></div>
                        <div className="export-card"><div className="export-card-icon"><Download size={18} /></div><div><strong>Ready to export?</strong><p>Download HTML with embedded images, plain text, Markdown, or JSON.</p></div><button aria-label="Open export options" className="export-card-arrow" onClick={() => setExportMenuPosition((value) => value === 'bottom' ? null : 'bottom')}><ChevronDown size={17} /></button></div>
                    </aside>}
                </div>

                <footer className="app-footer"><span>Word to HTML Studio <span className="footer-separator">·</span> Local-first conversion</span><span><span className="secure-icon"><Check size={11} /></span> Your document stays in this browser session</span></footer>
            </main>
            {showPaste && <PasteModal onClose={() => setShowPaste(false)} onConvert={importPastedHtml} />}
            {busy && <div className="busy-overlay"><div className="busy-card"><LoaderCircle size={24} className="spin" /><strong>Processing document</strong><span>Converting content and preparing images…</span></div></div>}
        </div>
    );
}

export default App;
