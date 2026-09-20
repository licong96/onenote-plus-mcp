import { describe, it, expect } from 'vitest';
import {
  htmlToMarkdown,
  htmlToOneNotePage,
  markdownToHtmlFragment,
  markdownToOneNoteHtml,
} from '@/markdown.js';

describe('markdownToOneNoteHtml', () => {
  it('wraps converted markdown in a full HTML document with the title', () => {
    const html = markdownToOneNoteHtml('# Hello\n\nA **bold** thing.', 'My Page');
    expect(html).toContain('<!DOCTYPE html>');
    expect(html).toContain('<title>My Page</title>');
    expect(html).toContain('<h1>Hello</h1>');
    expect(html).toContain('<strong>bold</strong>');
    // OneNote requires the XHTML namespace on the root element.
    expect(html).toContain('<html xmlns="http://www.w3.org/1999/xhtml">');
  });

  it('escapes HTML entities in the title', () => {
    const html = markdownToOneNoteHtml('body', 'A & B <c>');
    expect(html).toContain('<title>A &amp; B &lt;c&gt;</title>');
  });
});

describe('htmlToOneNotePage', () => {
  it('wraps fragment HTML in a OneNote document', () => {
    const out = htmlToOneNotePage('<p>hi</p>', 'T');
    expect(out).toContain('<title>T</title>');
    expect(out).toContain('<p>hi</p>');
    expect(out).toContain('<html xmlns="http://www.w3.org/1999/xhtml">');
  });

  it('passes through full HTML documents unchanged', () => {
    const input = '<html><head><title>X</title></head><body><p>x</p></body></html>';
    expect(htmlToOneNotePage(input, 'ignored')).toBe(input);
  });
});

describe('markdownToHtmlFragment', () => {
  it('returns HTML without a document wrapper', () => {
    const html = markdownToHtmlFragment('## Hello\n\n- one\n- two');
    expect(html).toContain('<h2>Hello</h2>');
    expect(html).toContain('<li>one</li>');
    expect(html).not.toContain('<html');
    expect(html).not.toContain('<body');
  });

  it('trims trailing whitespace that marked appends', () => {
    const html = markdownToHtmlFragment('Hello');
    expect(html).toBe('<p>Hello</p>');
  });
});

describe('htmlToMarkdown', () => {
  it('round-trips a simple structure', () => {
    const md = htmlToMarkdown('<h1>Heading</h1><p>Hello <em>world</em>.</p>');
    expect(md).toContain('# Heading');
    expect(md).toContain('Hello _world_.');
  });

  it('renders code blocks as fenced', () => {
    const md = htmlToMarkdown('<pre><code>const x = 1;</code></pre>');
    expect(md).toContain('```');
    expect(md).toContain('const x = 1;');
  });
});
