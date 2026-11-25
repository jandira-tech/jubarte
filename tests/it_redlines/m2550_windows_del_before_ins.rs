// SPDX-FileCopyrightText: 2026 Jandira Technologies, LLC
//
// SPDX-License-Identifier: Apache-2.0

//! #2550 — `--mode word-windows` writes a replacement's deletion before its
//! insertion, as Windows Word 2024 does: inside a paragraph, between the
//! paragraphs of a replaced block, and in the header, footer and footnote
//! stories. `--mode word` keeps Mac Word 16's order, the insertion first,
//! and so does `--mode powertools`.

use std::io::{Cursor, Read};

use jubarte::comparer::{CompareMode, WmlComparerSettings};
use jubarte::document_comparer::compare_documents_with_settings;
use jubarte::namespaces::W;
use jubarte::xmllinq::{Dom, NodeId, XName};

use crate::common::docx::{self, Part, W_NS, para};

const OD_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEADER: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
const FOOTER: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const FOOTNOTES: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";
const STYLES: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

/// The redline of `original` against `revised` in `mode`.
fn redline(original: &[u8], revised: &[u8], mode: CompareMode) -> Vec<u8> {
    compare_documents_with_settings(original, revised, &WmlComparerSettings::new(mode))
        .expect("the compare succeeds")
}

/// One XML part of a package.
fn part(docx: &[u8], name: &str) -> String {
    let mut zip = zip::ZipArchive::new(Cursor::new(docx.to_vec())).unwrap();
    let mut text = String::new();
    zip.by_name(name)
        .unwrap()
        .read_to_string(&mut text)
        .unwrap();
    text
}

/// `xml` parsed, with its root element.
fn parse(xml: &str) -> (Dom, NodeId) {
    let mut dom = Dom::new();
    let document = dom.parse_xdocument(xml);
    let root = dom.root(document).expect("a root element");
    (dom, root)
}

/// The `w:t` and `w:delText` under `node`, in order.
fn text(dom: &Dom, node: NodeId) -> String {
    let mut out = String::new();
    for n in dom.descendants(node, None) {
        if dom.name_is(n, &W::t()) || dom.name_is(n, &W::del_text()) {
            out.push_str(&dom.value(n));
        }
    }
    out
}

/// How a paragraph reads, without its properties: `d`, `i` or `x` (a plain
/// run) and the text of each child.
fn markup(dom: &Dom, paragraph: NodeId) -> Vec<(char, String)> {
    dom.elements(paragraph, None)
        .into_iter()
        .filter(|&c| !dom.name_is(c, &W::p_pr()))
        .map(|c| {
            let kind = match dom.name(c) {
                Some(n) if n == W::ins() => 'i',
                Some(n) if n == W::del() => 'd',
                _ => 'x',
            };
            (kind, text(dom, c))
        })
        .collect()
}

/// `D(text)` for a paragraph whose mark is deleted and whose every revision
/// wrapper is a `w:del`, `I(text)` for one inserted likewise, and `?(text)`
/// for any other paragraph; each paragraph directly under `container`.
fn blocks(dom: &Dom, container: NodeId) -> Vec<String> {
    dom.elements(container, Some(&W::p()))
        .into_iter()
        .map(|p| {
            let mark = dom
                .element(p, &W::p_pr())
                .and_then(|ppr| dom.element(ppr, &W::r_pr()));
            let mark_has = |rev: &XName| mark.is_some_and(|m| dom.element(m, rev).is_some());
            let kinds = markup(dom, p);
            let label = match (mark_has(&W::ins()), mark_has(&W::del())) {
                (true, false) if !kinds.is_empty() && kinds.iter().all(|&(k, _)| k == 'i') => 'I',
                (false, true) if !kinds.is_empty() && kinds.iter().all(|&(k, _)| k == 'd') => 'D',
                _ => '?',
            };
            format!("{label}({})", text(dom, p))
        })
        .collect()
}

/// The story's text with every revision accepted, and with every one
/// rejected.
fn views(dom: &Dom, root: NodeId) -> (String, String) {
    let (mut accepted, mut rejected) = (String::new(), String::new());
    for n in dom.descendants(root, None) {
        let is_text = dom.name_is(n, &W::t());
        let is_deleted = dom.name_is(n, &W::del_text());
        if !is_text && !is_deleted {
            continue;
        }
        let under = |name: &XName| !dom.ancestors_and_self(n, Some(name)).is_empty();
        if is_text && !under(&W::del()) {
            accepted.push_str(&dom.value(n));
        }
        if is_deleted || (is_text && !under(&W::ins())) {
            rejected.push_str(&dom.value(n));
        }
    }
    (accepted, rejected)
}

/// `Some(true)` for a paragraph whose mark is inserted and whose every run
/// sits directly under a `w:ins`, `Some(false)` for one deleted likewise,
/// and `None` for anything else.
fn whole(dom: &Dom, node: NodeId) -> Option<bool> {
    if !dom.name_is(node, &W::p()) {
        return None;
    }
    let mark = dom
        .element(node, &W::p_pr())
        .and_then(|ppr| dom.element(ppr, &W::r_pr()));
    let mark_has = |rev: &XName| mark.is_some_and(|m| dom.element(m, rev).is_some());
    let (inserted, deleted) = (mark_has(&W::ins()), mark_has(&W::del()));
    let (kind, wrapper) = match (inserted, deleted) {
        (true, false) => (true, W::ins()),
        (false, true) => (false, W::del()),
        _ => return None,
    };
    let whole = dom.descendants(node, Some(&W::r())).into_iter().all(|r| {
        dom.parent(r)
            .is_some_and(|parent| dom.name_is(parent, &wrapper))
    });
    whole.then_some(kind)
}

/// Nothing in the story reads an insertion ahead of a deletion it sits next
/// to: no run of adjacent `w:ins` / `w:del` wrappers (property containers
/// aside), and no run of adjacent whole-inserted and whole-deleted paragraphs
/// in a block container.
fn assert_deletions_first(xml: &str, label: &str) {
    let (dom, root) = parse(xml);
    let containers = [
        W::body(),
        W::tc(),
        W::footnote(),
        W::endnote(),
        W::hdr(),
        W::ftr(),
    ];
    let property_bag = [W::r_pr(), W::p_pr(), W::tr_pr()];
    for element in dom.descendants_and_self(root, None) {
        let mut inserted = false;
        if !property_bag.iter().any(|name| dom.name_is(element, name)) {
            for child in dom.elements(element, None) {
                if dom.name_is(child, &W::ins()) {
                    inserted = true;
                } else if dom.name_is(child, &W::del()) {
                    assert!(
                        !inserted,
                        "{label}: a deletion follows an insertion in {}",
                        dom.serialize_element(element)
                    );
                } else {
                    inserted = false;
                }
            }
        }
        if !containers.iter().any(|name| dom.name_is(element, name)) {
            continue;
        }
        let mut inserted = false;
        for child in dom.nodes(element) {
            match whole(&dom, child) {
                Some(true) => inserted = true,
                Some(false) => assert!(
                    !inserted,
                    "{label}: a deleted paragraph follows an inserted one in {}",
                    text(&dom, element)
                ),
                None => inserted = false,
            }
        }
    }
}

/// "Alpha beta gamma" against "Alpha delta gamma": one word replaced.
fn replaced_word(mode: CompareMode) -> (Dom, NodeId, Vec<u8>) {
    let original = docx::docx(&para("Alpha beta gamma"));
    let revised = docx::docx(&para("Alpha delta gamma"));
    let out = redline(&original, &revised, mode);
    let (dom, root) = parse(&part(&out, "word/document.xml"));
    (dom, root, out)
}

fn body_paragraphs(dom: &Dom, root: NodeId) -> Vec<NodeId> {
    dom.elements(dom.element(root, &W::body()).unwrap(), Some(&W::p()))
}

#[test]
fn a_replaced_word_is_deleted_then_inserted_in_word_windows() {
    let (dom, root, out) = replaced_word(CompareMode::WordWindows);
    assert_eq!(
        markup(&dom, body_paragraphs(&dom, root)[0]),
        [
            ('x', "Alpha ".to_string()),
            ('d', "beta".to_string()),
            ('i', "delta".to_string()),
            ('x', " gamma".to_string()),
        ]
    );
    assert_deletions_first(&part(&out, "word/document.xml"), "word-windows body");
    crate::common::validity::assert_word_valid_package(&out);
}

#[test]
fn the_same_word_is_still_inserted_then_deleted_in_word_and_powertools() {
    // Mac Word 16, the `word` preset's reference.
    let (dom, root, _) = replaced_word(CompareMode::Word);
    let word = markup(&dom, body_paragraphs(&dom, root)[0]);
    assert_eq!(
        word,
        [
            ('x', "Alpha ".to_string()),
            ('i', "delta".to_string()),
            ('d', "beta".to_string()),
            ('x', " gamma".to_string()),
        ]
    );
    // PowerTools-faithful keeps the insertion-first pass as well.
    let (dom, root, _) = replaced_word(CompareMode::PowerTools);
    assert_eq!(
        markup(&dom, body_paragraphs(&dom, root)[0]),
        [
            ('x', "Alpha ".to_string()),
            ('i', "delta".to_string()),
            ('d', "beta".to_string()),
            ('x', " gamma".to_string()),
        ]
    );
}

#[test]
fn the_two_orders_accept_and_reject_to_the_same_text() {
    let (word, word_root, _) = replaced_word(CompareMode::Word);
    let (windows, windows_root, _) = replaced_word(CompareMode::WordWindows);
    assert_eq!(views(&word, word_root), views(&windows, windows_root));
    assert_eq!(
        views(&windows, windows_root),
        (
            "Alpha delta gamma".to_string(),
            "Alpha beta gamma".to_string()
        )
    );
}

/// Two paragraphs replaced by two others, between two the revision keeps.
fn replaced_paragraphs(mode: CompareMode) -> (Dom, NodeId, Vec<u8>) {
    let original = docx::docx(&format!(
        "{}{}{}",
        para("First old title"),
        para("Second old title"),
        para("Tail")
    ));
    let revised = docx::docx(&format!(
        "{}{}{}",
        para("Brand new heading"),
        para("Another brand new heading"),
        para("Tail")
    ));
    let out = redline(&original, &revised, mode);
    let (dom, root) = parse(&part(&out, "word/document.xml"));
    (dom, root, out)
}

#[test]
fn replaced_paragraphs_read_deletions_first_in_word_windows() {
    let (dom, root, out) = replaced_paragraphs(CompareMode::WordWindows);
    let body = dom.element(root, &W::body()).unwrap();
    assert_eq!(
        blocks(&dom, body),
        [
            "D(First old title)",
            "D(Second old title)",
            "I(Brand new heading)",
            "I(Another brand new heading)",
            "?(Tail)",
        ]
    );
    assert_deletions_first(&part(&out, "word/document.xml"), "word-windows body");
    crate::common::validity::assert_word_valid_package(&out);
}

#[test]
fn replaced_paragraphs_keep_insertions_first_in_word() {
    let (word, word_root, _) = replaced_paragraphs(CompareMode::Word);
    assert_eq!(
        blocks(&word, word.element(word_root, &W::body()).unwrap()),
        [
            "I(Brand new heading)",
            "I(Another brand new heading)",
            "D(First old title)",
            "D(Second old title)",
            "?(Tail)",
        ]
    );
    let (windows, windows_root, _) = replaced_paragraphs(CompareMode::WordWindows);
    assert_eq!(views(&word, word_root), views(&windows, windows_root));
}

/// A document whose default header or footer says `story_text`.
fn with_story(kind: &str, tag: &str, content_type: &str, story_text: &str) -> Vec<u8> {
    let story = format!(
        r#"<w:{tag} xmlns:w="{W_NS}"><w:p><w:r><w:t xml:space="preserve">{story_text}</w:t></w:r></w:p></w:{tag}>"#
    );
    docx::docx_with_sect(
        &para("shared body"),
        &[Part {
            name: &format!("word/{kind}1.xml"),
            content_type,
            rel_type: &format!("{OD_REL}/{kind}"),
            xml: &story,
        }],
        &format!(r#"<w:{kind}Reference w:type="default" r:id="rIdX0"/>"#),
    )
}

#[test]
fn a_header_or_footer_replacement_is_deleted_first_in_word_windows() {
    for (kind, tag, content_type) in [("header", "hdr", HEADER), ("footer", "ftr", FOOTER)] {
        let original = with_story(kind, tag, content_type, "Header alpha text");
        let revised = with_story(kind, tag, content_type, "Header bravo text");
        let name = format!("word/{kind}1.xml");

        let out = redline(&original, &revised, CompareMode::WordWindows);
        let (dom, root) = parse(&part(&out, &name));
        let paragraphs = dom.descendants(root, Some(&W::p()));
        assert_eq!(paragraphs.len(), 1, "{kind}");
        assert_eq!(
            markup(&dom, paragraphs[0]),
            [
                ('x', "Header ".to_string()),
                ('d', "alpha".to_string()),
                ('i', "bravo".to_string()),
                ('x', " text".to_string()),
            ],
            "{kind}"
        );
        assert_deletions_first(&part(&out, &name), &name);
        crate::common::validity::assert_word_valid_package(&out);

        // The same story in `word` keeps Mac Word's order.
        let out = redline(&original, &revised, CompareMode::Word);
        let (dom, root) = parse(&part(&out, &name));
        let word = markup(&dom, dom.descendants(root, Some(&W::p()))[0]);
        assert_eq!(
            word,
            [
                ('x', "Header ".to_string()),
                ('i', "bravo".to_string()),
                ('d', "alpha".to_string()),
                ('x', " text".to_string()),
            ],
            "{kind}"
        );
    }
}

/// A document with one footnote whose definition holds `paragraphs`.
fn with_footnote(paragraphs: &[&str]) -> Vec<u8> {
    let body = r#"<w:p><w:r><w:t xml:space="preserve">Shared body</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>"#;
    let definitions: String = paragraphs
        .iter()
        .map(|text| {
            format!(
                r#"<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#
            )
        })
        .collect();
    let notes = format!(
        r#"<w:footnotes xmlns:w="{W_NS}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">{definitions}</w:footnote></w:footnotes>"#
    );
    let styles = format!(
        r#"<w:styles xmlns:w="{W_NS}"><w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/></w:style><w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/></w:style></w:styles>"#
    );
    docx::docx_with(
        body,
        &[
            Part {
                name: "word/footnotes.xml",
                content_type: FOOTNOTES,
                rel_type: &format!("{OD_REL}/footnotes"),
                xml: &notes,
            },
            Part {
                name: "word/styles.xml",
                content_type: STYLES,
                rel_type: &format!("{OD_REL}/styles"),
                xml: &styles,
            },
        ],
    )
}

/// The paragraphs of the redline's one content footnote, `w:id="1"` (the
/// separators are `w:footnote` elements too).
fn footnote_paragraphs(out: &[u8]) -> (Dom, Vec<NodeId>) {
    let (dom, root) = parse(&part(out, "word/footnotes.xml"));
    let definition = dom
        .elements(root, Some(&W::footnote()))
        .into_iter()
        .find(|&n| dom.attribute(n, &W::id()) == Some("1"))
        .expect("the content footnote");
    let paragraphs = dom.elements(definition, Some(&W::p()));
    (dom, paragraphs)
}

/// The notes are compared by their own path, which has no insertion-first
/// pass in any mode: a replacement there already reads deletion first. The
/// contract is the same in `word-windows`, where it must stay whatever the
/// body's order comes to.
#[test]
fn footnotes_read_deletions_first_in_word_windows() {
    let original = with_footnote(&["Alpha beta gamma"]);
    let revised = with_footnote(&["Alpha delta gamma"]);
    let out = redline(&original, &revised, CompareMode::WordWindows);
    let (dom, paragraphs) = footnote_paragraphs(&out);
    assert_eq!(paragraphs.len(), 1, "one paragraph in the note");
    assert_eq!(
        markup(&dom, paragraphs[0]),
        [
            ('x', "".to_string()),
            ('x', "Alpha ".to_string()),
            ('d', "beta".to_string()),
            ('i', "delta".to_string()),
            ('x', " gamma".to_string()),
        ]
    );
    assert_deletions_first(&part(&out, "word/footnotes.xml"), "word-windows footnotes");
    crate::common::validity::assert_word_valid_package(&out);

    // Paragraphs of the note replaced whole.
    let original = with_footnote(&["Old one note", "Old two note"]);
    let revised = with_footnote(&["Brand new heading", "Another brand new heading"]);
    let out = redline(&original, &revised, CompareMode::WordWindows);
    let (dom, paragraphs) = footnote_paragraphs(&out);
    let texts: Vec<String> = paragraphs.iter().map(|&p| text(&dom, p)).collect();
    assert_eq!(
        texts,
        [
            "Old one note",
            "Old two note",
            "Brand new heading",
            "Another brand new heading",
        ]
    );
    assert_deletions_first(&part(&out, "word/footnotes.xml"), "word-windows footnotes");
    crate::common::validity::assert_word_valid_package(&out);
}
