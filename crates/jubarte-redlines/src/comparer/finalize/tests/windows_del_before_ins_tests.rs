// SPDX-FileCopyrightText: 2026 Jandira Technologies, LLC
//
// SPDX-License-Identifier: Apache-2.0

//! #2550 — Windows Word 2024 writes a replacement's deletion before its
//! insertion. Mac Word 16, the `word` preset's reference, writes the
//! insertion first ([`reorder_replacements_ins_before_del`]); `word-windows`
//! runs these two passes in its place, inside each paragraph and between
//! the paragraphs of a replaced block.
//!
//! #2555 — the hardening: the inline pass partitions within each
//! (author, date) group and orders an all-one-kind hyperlink with its
//! kind; the paragraph pass classifies runs behind containers, walks the
//! text-box, content-control and comment bodies, and leaves section-break
//! paragraphs and the story's final pilcrow where they stand.

use super::*;

const BY: &str = r#"w:author="Editor" w:date="2026-01-01T00:00:00Z""#;
const OTHER: &str = r#"w:author="Other" w:date="2026-02-02T00:00:00Z""#;

/// A `w:del` (`del`) or `w:ins` (`ins`) wrapper around one run of `text`,
/// stamped `attrs`.
fn wrapper_by(kind: &str, id: u32, text: &str, attrs: &str) -> String {
    let leaf = if kind == "del" { "delText" } else { "t" };
    format!(
        r#"<w:{kind} w:id="{id}" {attrs}><w:r><w:{leaf} xml:space="preserve">{text}</w:{leaf}></w:r></w:{kind}>"#
    )
}

/// A `w:del` (`del`) or `w:ins` (`ins`) wrapper around one run of `text`.
fn wrapper(kind: &str, id: u32, text: &str) -> String {
    wrapper_by(kind, id, text, BY)
}

/// A paragraph whose mark carries `mark` (`ins` or `del`), around `body`.
fn marked(mark: &str, body: &str) -> String {
    format!(r#"<w:p><w:pPr><w:rPr><w:{mark} w:id="1" {BY}/></w:rPr></w:pPr>{body}</w:p>"#)
}

/// A wholly inserted (`ins`) or wholly deleted (`del`) paragraph on `text`.
fn whole(kind: &str, id: u32, text: &str) -> String {
    marked(kind, &wrapper(kind, id, text))
}

/// An unchanged paragraph on `text`.
fn plain(text: &str) -> String {
    format!(r#"<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#)
}

/// A run holding one field char of the given type.
fn field_char(kind: &str) -> String {
    format!(r#"<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>"#)
}

/// A document whose body holds `fragment`, parsed.
fn document(fragment: &str) -> (Dom, NodeId) {
    let mut dom = Dom::new();
    let doc = dom.parse_xdocument(&format!(
        r#"<w:document xmlns:w="{}"><w:body>{fragment}<w:sectPr/></w:body></w:document>"#,
        W::URI
    ));
    let root = dom.root(doc).unwrap();
    (dom, root)
}

/// All the text under `node`, in document order.
fn text(dom: &Dom, node: NodeId) -> String {
    dom.descendant_nodes(node)
        .into_iter()
        .filter_map(|n| dom.text_value(n))
        .collect()
}

/// A paragraph holding `runs`, as the first paragraph of a document.
fn paragraph(runs: &str) -> String {
    format!("<w:p>{runs}</w:p>")
}

/// The first paragraph's revision wrappers in order: `d` for a `w:del`, `i`
/// for a `w:ins`, each with its text.
fn wrappers(dom: &Dom, root: NodeId) -> Vec<(char, String)> {
    let p = dom.descendants(root, Some(&W::p()))[0];
    dom.elements(p, None)
        .into_iter()
        .filter_map(|c| {
            let kind = match dom.name(c) {
                Some(n) if n == W::del() => 'd',
                Some(n) if n == W::ins() => 'i',
                _ => return None,
            };
            Some((kind, text(dom, c)))
        })
        .collect()
}

/// Local names of the element children of `node`.
fn children(dom: &Dom, node: NodeId) -> Vec<String> {
    dom.elements(node, None)
        .into_iter()
        .filter_map(|c| dom.name(c).map(|n| n.local_name().to_string()))
        .collect()
}

fn body(dom: &Dom, root: NodeId) -> NodeId {
    dom.element(root, &W::body()).unwrap()
}

/// The text of each paragraph directly under `container`, in order.
fn texts(dom: &Dom, container: NodeId) -> Vec<String> {
    dom.elements(container, Some(&W::p()))
        .into_iter()
        .map(|p| text(dom, p))
        .collect()
}

/// The text that accepting (`true`) or rejecting the revisions reads.
fn view(dom: &Dom, root: NodeId, accepted: bool) -> String {
    let (leaf, hidden) = if accepted {
        (W::t(), W::del())
    } else {
        (W::del_text(), W::ins())
    };
    dom.descendants(root, Some(&leaf))
        .into_iter()
        .filter(|&n| dom.ancestors_and_self(n, Some(&hidden)).is_empty())
        .map(|n| dom.value(n))
        .collect()
}

#[test]
fn an_insertion_before_a_deletion_becomes_a_deletion_before_an_insertion() {
    let (mut dom, root) = document(&paragraph(&format!(
        "{}{}",
        wrapper("ins", 1, "new"),
        wrapper("del", 2, "old")
    )));
    let (accepted, rejected) = (view(&dom, root, true), view(&dom, root, false));
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(
        wrappers(&dom, root),
        [('d', "old".to_string()), ('i', "new".to_string())]
    );
    assert_eq!(
        (view(&dom, root, true), view(&dom, root, false)),
        (accepted, rejected)
    );
}

#[test]
fn a_mixed_run_keeps_the_order_inside_each_kind() {
    // ins, ins, del, ins, del -> del, del, ins, ins, ins.
    let (mut dom, root) = document(&paragraph(&format!(
        "{}{}{}{}{}",
        wrapper("ins", 1, "i1"),
        wrapper("ins", 2, "i2"),
        wrapper("del", 3, "d1"),
        wrapper("ins", 4, "i3"),
        wrapper("del", 5, "d2")
    )));
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(
        wrappers(&dom, root),
        [
            ('d', "d1".to_string()),
            ('d', "d2".to_string()),
            ('i', "i1".to_string()),
            ('i', "i2".to_string()),
            ('i', "i3".to_string())
        ]
    );
}

/// A run of wrappers in which one holds part of a field.
#[test]
fn a_run_holding_half_a_field_stays_where_it_is() {
    let half = format!(r#"<w:ins w:id="1" {BY}>{}</w:ins>"#, field_char("begin"));
    let end_alone = format!(r#"<w:del w:id="2" {BY}>{}</w:del>"#, field_char("end"));
    for fragment in [
        format!("{half}{}", wrapper("del", 2, "old")),
        // One half-field wrapper holds the whole run, its whole wrappers
        // included.
        format!(
            "{half}{}{}",
            wrapper("ins", 3, "i1"),
            wrapper("del", 2, "old")
        ),
        format!(
            "{}{}{end_alone}",
            wrapper("ins", 1, "i1"),
            wrapper("del", 3, "d1")
        ),
    ] {
        let (mut dom, root) = document(&paragraph(&fragment));
        let before = dom.serialize_element(root);
        reorder_deletions_before_insertions(&mut dom, root);
        assert_eq!(dom.serialize_element(root), before, "{fragment}");
    }
    // A wrapper whose own field is complete moves.
    let complete = format!(
        r#"<w:ins w:id="1" {BY}>{}<w:r><w:instrText> PAGE </w:instrText></w:r>{}<w:r><w:t>7</w:t></w:r>{}</w:ins>"#,
        field_char("begin"),
        field_char("separate"),
        field_char("end")
    );
    let (mut dom, root) = document(&paragraph(&format!(
        "{complete}{}",
        wrapper("del", 2, "old")
    )));
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(
        wrappers(&dom, root),
        [('d', "old".to_string()), ('i', " PAGE 7".to_string())]
    );
    // A field char outside the wrappers' run is no concern of the run.
    let (mut dom, root) = document(&paragraph(&format!(
        "{}{}{}",
        wrapper("ins", 5, "i1"),
        wrapper("del", 4, "d1"),
        field_char("end")
    )));
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(
        wrappers(&dom, root),
        [('d', "d1".to_string()), ('i', "i1".to_string())]
    );
    let p = dom.descendants(root, Some(&W::p()))[0];
    assert_eq!(children(&dom, p), ["del", "ins", "r"]);
}

/// The paragraph-mark, run-property and row-property revisions sit in
/// `w:rPr` / `w:trPr`, where `w:ins` before `w:del` is the schema's order.
#[test]
fn property_containers_keep_their_revision_order() {
    let revision = |kind: &str, id: u32| format!(r#"<w:{kind} w:id="{id}" {BY}/>"#);
    let (mut dom, root) = document(&format!(
        concat!(
            "<w:p><w:pPr><w:rPr>{}{}</w:rPr></w:pPr>{}{}</w:p>",
            "<w:tbl><w:tr><w:trPr>{}{}</w:trPr></w:tr></w:tbl>",
            "<w:p><w:r><w:rPr>{}{}</w:rPr><w:t>kept</w:t></w:r></w:p>"
        ),
        revision("ins", 1),
        revision("del", 2),
        wrapper("ins", 3, "new"),
        wrapper("del", 4, "old"),
        revision("ins", 5),
        revision("del", 6),
        revision("ins", 7),
        revision("del", 8),
    ));
    reorder_deletions_before_insertions(&mut dom, root);
    let runs = dom.descendants(root, Some(&W::r_pr()));
    assert_eq!(runs.len(), 2, "the paragraph mark's and the run's");
    for rpr in runs {
        assert_eq!(children(&dom, rpr), ["ins", "del"]);
    }
    assert_eq!(
        children(&dom, dom.descendants(root, Some(&W::tr_pr()))[0]),
        ["ins", "del"]
    );
    // The paragraph's own wrappers still trade places.
    let paragraphs = dom.descendants(root, Some(&W::p()));
    assert_eq!(children(&dom, paragraphs[0]), ["pPr", "del", "ins"]);
    assert_eq!(children(&dom, paragraphs[1]), ["r"]);
}

/// `d` for each `w:del` and `i` for each `w:ins` child of the paragraph, in
/// order; other children are left out.
fn kinds(dom: &Dom, p: NodeId) -> Vec<char> {
    dom.elements(p, None)
        .into_iter()
        .filter_map(|c| match dom.name(c) {
            Some(n) if n == W::del() => Some('d'),
            Some(n) if n == W::ins() => Some('i'),
            _ => None,
        })
        .collect()
}

#[test]
fn the_wrapper_pass_reaches_every_paragraph_and_is_idempotent() {
    let kept = r#"<w:r><w:t>kept</w:t></w:r>"#;
    let half_field = format!(r#"<w:del w:id="4" {BY}>{}</w:del>"#, field_char("end"));
    // What the first paragraph's wrappers come to, and what builds them.
    let cases = [
        // The last wrappers of the paragraph, with no sibling after them.
        (
            format!("{}{}", wrapper("ins", 1, "i1"), wrapper("del", 2, "d1")),
            "di",
        ),
        (
            format!(
                "{}{}{}",
                wrapper("ins", 1, "i1"),
                wrapper("del", 2, "d1"),
                wrapper("ins", 3, "i2")
            ),
            "dii",
        ),
        (
            format!(
                "{}{}{}",
                wrapper("del", 1, "d0"),
                wrapper("ins", 2, "i1"),
                wrapper("del", 3, "d1")
            ),
            "ddi",
        ),
        // An unchanged run, then a field char in a run of its own: both
        // follow the wrappers' run.
        (
            format!(
                "{}{}{kept}{}",
                wrapper("ins", 1, "i1"),
                wrapper("del", 2, "d1"),
                field_char("end")
            ),
            "di",
        ),
        // A half field inside a wrapper holds the run.
        (
            format!(
                "{}{}{}{}",
                wrapper("ins", 1, "i1"),
                wrapper("del", 2, "d1"),
                wrapper("ins", 3, "i2"),
                half_field
            ),
            "idid",
        ),
        // An unchanged run between two runs: each is a run of its own, the
        // first already in order and the second not.
        (
            format!(
                "{}{}{kept}{}{}",
                wrapper("del", 1, "d1"),
                wrapper("ins", 2, "i1"),
                wrapper("ins", 3, "i2"),
                wrapper("del", 4, "d2")
            ),
            "didi",
        ),
    ];
    for (runs, expected) in cases {
        let (mut dom, root) = document(&paragraph(&runs));
        reorder_deletions_before_insertions(&mut dom, root);
        assert_eq!(
            kinds(&dom, dom.descendants(root, Some(&W::p()))[0]),
            expected.chars().collect::<Vec<_>>(),
            "{runs}"
        );
        let once = dom.serialize_element(root);
        reorder_deletions_before_insertions(&mut dom, root);
        assert_eq!(dom.serialize_element(root), once, "{runs}");
    }
    // Inside a table cell's paragraph and in a run of a text box.
    let (mut dom, root) = document(&format!(
        concat!(
            "<w:tbl><w:tr><w:tc><w:p>{}{}</w:p></w:tc></w:tr></w:tbl>",
            "<w:p>{}<w:r><mc:AlternateContent xmlns:mc=\"{}\"><w:pict><v:textbox xmlns:v=\"x\"><w:txbxContent><w:p>{}{}</w:p></w:txbxContent></v:textbox></w:pict></mc:AlternateContent></w:r></w:p>"
        ),
        wrapper("ins", 1, "i1"),
        wrapper("del", 2, "d1"),
        kept,
        crate::namespaces::MC::URI,
        wrapper("ins", 3, "i2"),
        wrapper("del", 4, "d2")
    ));
    reorder_deletions_before_insertions(&mut dom, root);
    let paragraphs = dom.descendants(root, Some(&W::p()));
    assert_eq!(paragraphs.len(), 3, "the cell's, the body's and the box's");
    assert_eq!(kinds(&dom, paragraphs[0]), ['d', 'i']);
    assert_eq!(kinds(&dom, paragraphs[1]), Vec::<char>::new());
    assert_eq!(kinds(&dom, paragraphs[2]), ['d', 'i']);
    let once = dom.serialize_element(root);
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(dom.serialize_element(root), once);
}

#[test]
fn inserted_paragraphs_follow_the_deleted_ones() {
    let (mut dom, root) = document(&format!(
        "{}{}{}",
        whole("ins", 1, "i1"),
        whole("ins", 2, "i2"),
        whole("del", 3, "d1")
    ));
    reorder_deleted_paragraphs_before_inserted(&mut dom, root);
    assert_eq!(texts(&dom, body(&dom, root)), ["d1", "i1", "i2"]);
}

#[test]
fn an_unchanged_paragraph_between_them_leaves_both_where_they_are() {
    let (mut dom, root) = document(&format!(
        "{}{}{}",
        whole("ins", 1, "i1"),
        plain("kept"),
        whole("del", 2, "d1")
    ));
    let before = dom.serialize_element(root);
    reorder_deleted_paragraphs_before_inserted(&mut dom, root);
    assert_eq!(dom.serialize_element(root), before);
    assert_eq!(texts(&dom, body(&dom, root)), ["i1", "kept", "d1"]);
}

/// A paragraph whose mark says one thing and whose runs say another breaks
/// the run: it belongs to neither side.
#[test]
fn a_revised_mark_alone_does_not_make_a_paragraph_whole() {
    let both_marks = format!(
        r#"<w:p><w:pPr><w:rPr><w:ins w:id="1" {BY}/><w:del w:id="2" {BY}/></w:rPr></w:pPr>{}</w:p>"#,
        wrapper("ins", 3, "half")
    );
    let halves = [
        // An inserted mark around one plain run.
        marked("ins", r#"<w:r><w:t>half</w:t></w:r>"#),
        // A deleted mark around one inserted run.
        marked("del", &wrapper("ins", 3, "half")),
        // A deleted mark around a deleted run and a plain one.
        marked(
            "del",
            &format!(
                r#"<w:r><w:t>plain </w:t></w:r>{}"#,
                wrapper("del", 4, "half")
            ),
        ),
        both_marks,
        // An inserted mark around a deleted hyperlink: the run's first
        // wrapper is the wrong kind, so the paragraph is neither.
        marked(
            "ins",
            r#"<w:del w:id="3" w:author="Editor" w:date="2026-01-01T00:00:00Z"><w:hyperlink w:anchor="x"><w:r><w:t>half</w:t></w:r></w:hyperlink></w:del>"#,
        ),
    ];
    for half in halves {
        let (mut dom, root) = document(&format!(
            "{}{half}{}",
            whole("ins", 5, "i1"),
            whole("del", 6, "d1")
        ));
        let before = dom.serialize_element(root);
        reorder_deleted_paragraphs_before_inserted(&mut dom, root);
        assert_eq!(dom.serialize_element(root), before, "{half}");
        // Last of the run, it is still neither.
        let (mut dom, root) = document(&format!("{}{half}", whole("del", 6, "d1")));
        let before = dom.serialize_element(root);
        reorder_deleted_paragraphs_before_inserted(&mut dom, root);
        assert_eq!(dom.serialize_element(root), before, "{half}");
    }
}

#[test]
fn an_empty_paragraph_counts_by_its_mark() {
    for mark in ["ins", "del"] {
        let other = ["del", "ins"].into_iter().find(|&k| k != mark).unwrap();
        let (mut dom, root) = document(&format!("{}{}", marked(mark, ""), whole(other, 2, "x1")));
        reorder_deleted_paragraphs_before_inserted(&mut dom, root);
        let expected = if mark == "del" {
            vec![String::new(), "x1".to_string()]
        } else {
            vec!["x1".to_string(), String::new()]
        };
        assert_eq!(texts(&dom, body(&dom, root)), expected, "{mark}");
    }
}

#[test]
fn the_paragraph_pass_keeps_each_kinds_order_reads_the_same_and_is_idempotent() {
    let (mut dom, root) = document(&format!(
        "{}{}{}{}{}{}",
        plain("kept"),
        whole("ins", 1, "i1"),
        whole("del", 2, "d1"),
        plain("between"),
        whole("ins", 3, "i2"),
        whole("del", 4, "d2")
    ));
    let views = (view(&dom, root, true), view(&dom, root, false));
    reorder_deleted_paragraphs_before_inserted(&mut dom, root);
    assert_eq!(
        texts(&dom, body(&dom, root)),
        ["kept", "d1", "i1", "between", "d2", "i2"]
    );
    assert_eq!((view(&dom, root, true), view(&dom, root, false)), views);
    let once = dom.serialize_element(root);
    reorder_deleted_paragraphs_before_inserted(&mut dom, root);
    assert_eq!(dom.serialize_element(root), once);

    // Deletions first already: nothing moves.
    let (mut dom, root) = document(&format!(
        "{}{}{}{}",
        whole("del", 1, "d1"),
        whole("del", 2, "d2"),
        plain("kept"),
        whole("ins", 3, "i1")
    ));
    let before = dom.serialize_element(root);
    reorder_deleted_paragraphs_before_inserted(&mut dom, root);
    assert_eq!(dom.serialize_element(root), before);
    assert_eq!(texts(&dom, body(&dom, root)), ["d1", "d2", "kept", "i1"]);
}

#[test]
fn every_block_container_orders_its_own_paragraphs() {
    let run = |first_id: u32| {
        format!(
            "{}{}{}",
            whole("ins", first_id, "i1"),
            whole("ins", first_id + 1, "i2"),
            whole("del", first_id + 2, "d1")
        )
    };
    let (mut dom, root) = document(&format!(
        concat!(
            "{}<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>",
            "<w:footnotes><w:footnote>{}</w:footnote></w:footnotes>",
            "<w:endnotes><w:endnote>{}</w:endnote></w:endnotes>",
            "<w:hdr>{}</w:hdr><w:ftr>{}</w:ftr>"
        ),
        run(10),
        run(20),
        run(30),
        run(40),
        run(50),
        run(60)
    ));
    reorder_deleted_paragraphs_before_inserted(&mut dom, root);
    assert_eq!(texts(&dom, body(&dom, root)), ["d1", "i1", "i2"]);
    for name in [W::tc(), W::footnote(), W::endnote(), W::hdr(), W::ftr()] {
        let container = dom.descendants(root, Some(&name))[0];
        assert_eq!(texts(&dom, container), ["d1", "i1", "i2"], "{name:?}");
    }
}

// ---- #2555: the inline pass partitions within each (author, date) group
// and orders an all-one-kind hyperlink with its kind ----

/// Two authors, each having replaced one word: `[del A][ins B][del C][ins D]`
/// with A/B by one author and C/D by the other. Each author's replacement
/// is already in Windows Word's order (deletion first), so nothing moves;
/// the blind whole-run partition would fuse them into
/// `[del A][del C][ins B][ins D]`.
#[test]
fn two_authors_replacements_stay_two_replacements() {
    let stable = format!(
        "{}{}{}{}",
        wrapper_by("del", 1, "A", BY),
        wrapper_by("ins", 2, "B", BY),
        wrapper_by("del", 3, "C", OTHER),
        wrapper_by("ins", 4, "D", OTHER),
    );
    let (mut dom, root) = document(&paragraph(&stable));
    let (accepted, rejected) = (view(&dom, root, true), view(&dom, root, false));
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(
        wrappers(&dom, root),
        [
            ('d', "A".to_string()),
            ('i', "B".to_string()),
            ('d', "C".to_string()),
            ('i', "D".to_string()),
        ]
    );
    assert_eq!(
        (view(&dom, root, true), view(&dom, root, false)),
        (accepted, rejected)
    );

    // Mac Word's order, two authors: each replacement turns on its own.
    let mac = format!(
        "{}{}{}{}",
        wrapper_by("ins", 1, "B", BY),
        wrapper_by("del", 2, "A", BY),
        wrapper_by("ins", 3, "D", OTHER),
        wrapper_by("del", 4, "C", OTHER),
    );
    let (mut dom, root) = document(&paragraph(&mac));
    let (accepted, rejected) = (view(&dom, root, true), view(&dom, root, false));
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(
        wrappers(&dom, root),
        [
            ('d', "A".to_string()),
            ('i', "B".to_string()),
            ('d', "C".to_string()),
            ('i', "D".to_string()),
        ]
    );
    assert_eq!(
        (view(&dom, root, true), view(&dom, root, false)),
        (accepted, rejected)
    );

    // Interleaved authors never trade: no group holds both kinds.
    let mixed = format!(
        "{}{}{}{}",
        wrapper_by("ins", 1, "B", BY),
        wrapper_by("del", 2, "C", OTHER),
        wrapper_by("ins", 3, "D", OTHER),
        wrapper_by("del", 4, "A", BY),
    );
    let (mut dom, root) = document(&paragraph(&mixed));
    let before = dom.serialize_element(root);
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(dom.serialize_element(root), before);

    // One author, two replacements: the run still fuses, as before.
    let (mut dom, root) = document(&paragraph(&format!(
        "{}{}{}{}",
        wrapper("ins", 1, "B"),
        wrapper("del", 2, "A"),
        wrapper("ins", 3, "D"),
        wrapper("del", 4, "C"),
    )));
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(
        wrappers(&dom, root),
        [
            ('d', "A".to_string()),
            ('d', "C".to_string()),
            ('i', "B".to_string()),
            ('i', "D".to_string()),
        ]
    );
}

/// A `w:hyperlink` whose element children are all one kind trades with that
/// kind: the insertion before a deleted link moves behind it, and the
/// deleted link before an insertion stays ahead of it.
#[test]
fn an_all_one_kind_hyperlink_orders_with_its_kind() {
    let deleted_link = format!(
        r#"<w:hyperlink w:anchor="x">{}</w:hyperlink>"#,
        wrapper("del", 1, "Click here now")
    );
    let inserted_link = format!(
        r#"<w:hyperlink w:anchor="x">{}</w:hyperlink>"#,
        wrapper("ins", 1, "Click here now")
    );
    // Mac Word's order: the inserted text, then the deleted link.
    let (mut dom, root) = document(&paragraph(&format!(
        "{}{deleted_link}",
        wrapper("ins", 2, "One two")
    )));
    let (accepted, rejected) = (view(&dom, root, true), view(&dom, root, false));
    reorder_deletions_before_insertions(&mut dom, root);
    let p = dom.descendants(root, Some(&W::p()))[0];
    assert_eq!(
        children(&dom, p),
        ["hyperlink", "ins"],
        "the deleted link precedes the insertion"
    );
    assert_eq!(
        (view(&dom, root, true), view(&dom, root, false)),
        (accepted, rejected)
    );

    // The mirror: a deleted wrapper before an inserted link.
    let (mut dom, root) = document(&paragraph(&format!(
        "{inserted_link}{}",
        wrapper("del", 2, "One two")
    )));
    let (accepted, rejected) = (view(&dom, root, true), view(&dom, root, false));
    reorder_deletions_before_insertions(&mut dom, root);
    let p = dom.descendants(root, Some(&W::p()))[0];
    assert_eq!(
        children(&dom, p),
        ["del", "hyperlink"],
        "the deletion precedes the inserted link"
    );
    assert_eq!(
        (view(&dom, root, true), view(&dom, root, false)),
        (accepted, rejected)
    );

    // A link of mixed kinds, or by another author, is a boundary: the
    // insertion ahead of it stays put.
    let mixed_link = format!(
        r#"<w:hyperlink w:anchor="x">{}{}</w:hyperlink>"#,
        wrapper("del", 1, "old "),
        wrapper("ins", 2, "new")
    );
    let (mut dom, root) = document(&paragraph(&format!(
        "{}{mixed_link}",
        wrapper("ins", 3, "One two")
    )));
    let before = dom.serialize_element(root);
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(dom.serialize_element(root), before);

    let other_link = format!(
        r#"<w:hyperlink w:anchor="x">{}</w:hyperlink>"#,
        wrapper_by("del", 1, "Click here now", OTHER)
    );
    let (mut dom, root) = document(&paragraph(&format!(
        "{}{other_link}",
        wrapper("ins", 2, "One two")
    )));
    let before = dom.serialize_element(root);
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(dom.serialize_element(root), before);

    // Idempotent on the trading shape.
    let (mut dom, root) = document(&paragraph(&format!(
        "{}{deleted_link}",
        wrapper("ins", 2, "One two")
    )));
    reorder_deletions_before_insertions(&mut dom, root);
    let once = dom.serialize_element(root);
    reorder_deletions_before_insertions(&mut dom, root);
    assert_eq!(dom.serialize_element(root), once);
}

// ---- #2555: the paragraph pass classifies runs behind containers ----

/// A whole-revised paragraph whose runs sit behind a `w:hyperlink`,
/// `w:sdt`, `w:smartTag` or `w:fldSimple` the wrapper holds trades with
/// the other side's blocks; before #2555 only a run directly under the
/// wrapper counted, and these paragraphs dropped out of the pass.
#[test]
fn runs_behind_containers_count_as_whole_revised() {
    let cases: [(&str, &str, &str); 5] = [
        // Word's own pre-hoist shape: the hyperlink inside the wrapper.
        (
            "ins",
            "half",
            r#"<w:ins w:id="3" w:author="Editor" w:date="2026-01-01T00:00:00Z"><w:hyperlink w:anchor="x"><w:r><w:t>half</w:t></w:r></w:hyperlink></w:ins>"#,
        ),
        // The valid shape: the wrapper inside the hyperlink.
        (
            "ins",
            "half",
            r#"<w:hyperlink w:anchor="x"><w:ins w:id="3" w:author="Editor" w:date="2026-01-01T00:00:00Z"><w:r><w:t>half</w:t></w:r></w:ins></w:hyperlink>"#,
        ),
        // A run-level content control the wrapper holds.
        (
            "ins",
            "half",
            r#"<w:ins w:id="3" w:author="Editor" w:date="2026-01-01T00:00:00Z"><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>half</w:t></w:r></w:sdtContent></w:sdt></w:ins>"#,
        ),
        // A smart tag the wrapper holds.
        (
            "ins",
            "half",
            r#"<w:ins w:id="3" w:author="Editor" w:date="2026-01-01T00:00:00Z"><w:smartTag w:element="x"><w:r><w:t>half</w:t></w:r></w:smartTag></w:ins>"#,
        ),
        // A simple field the wrapper holds.
        (
            "del",
            "7",
            r#"<w:del w:id="3" w:author="Editor" w:date="2026-01-01T00:00:00Z"><w:fldSimple w:instr=" PAGE "><w:r><w:delText>7</w:delText></w:r></w:fldSimple></w:del>"#,
        ),
    ];
    for (mark, own_text, runs) in cases {
        let other = ["ins", "del"].into_iter().find(|&k| k != mark).unwrap();
        let case = marked(mark, runs);
        // The container paragraph trades with an opposite block beside it,
        // whichever side makes the trade necessary.
        let fragment = if mark == "ins" {
            format!("{case}{}", whole(other, 5, "mate"))
        } else {
            format!("{}{case}", whole(other, 5, "mate"))
        };
        let (mut dom, root) = document(&fragment);
        let (accepted, rejected) = (view(&dom, root, true), view(&dom, root, false));
        reorder_deleted_paragraphs_before_inserted(&mut dom, root);
        assert_eq!(
            texts(&dom, body(&dom, root)),
            if mark == "ins" {
                vec!["mate".to_string(), own_text.to_string()]
            } else {
                vec![own_text.to_string(), "mate".to_string()]
            },
            "{case}"
        );
        assert_eq!(
            (view(&dom, root, true), view(&dom, root, false)),
            (accepted, rejected),
            "{case}"
        );
        // Idempotent.
        let once = dom.serialize_element(root);
        reorder_deleted_paragraphs_before_inserted(&mut dom, root);
        assert_eq!(dom.serialize_element(root), once, "{case}");
    }
}

// ---- #2555: the paragraph pass walks the text-box, content-control and
// comment bodies ----

/// A `w:txbxContent`, `w:sdtContent` and `w:comment` each order their own
/// whole-revised paragraphs, as `w:body` and the other containers do; a
/// `w:sdt` block between two body paragraphs ends the body's run.
#[test]
fn text_boxes_content_controls_and_comment_bodies_order_their_paragraphs() {
    let run = |first_id: u32| {
        format!(
            "{}{}{}",
            whole("ins", first_id, "i1"),
            whole("ins", first_id + 1, "i2"),
            whole("del", first_id + 2, "d1")
        )
    };
    let (mut dom, root) = document(&format!(
        concat!(
            r#"<w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="tb"><v:textbox><w:txbxContent>{}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"#,
            "<w:sdt><w:sdtPr/><w:sdtContent>{}</w:sdtContent></w:sdt>",
            "<w:comments><w:comment w:id=\"0\">{}</w:comment></w:comments>"
        ),
        run(10),
        run(20),
        run(30)
    ));
    reorder_deleted_paragraphs_before_inserted(&mut dom, root);
    for name in [W::name("txbxContent"), W::sdt_content(), W::comment()] {
        let container = dom.descendants(root, Some(&name))[0];
        assert_eq!(texts(&dom, container), ["d1", "i1", "i2"], "{name:?}");
    }
    let once = dom.serialize_element(root);
    reorder_deleted_paragraphs_before_inserted(&mut dom, root);
    assert_eq!(dom.serialize_element(root), once);
}

// ---- #2555: the block pass's guards ----

/// A paragraph whose `w:pPr` carries a `w:sectPr` ends a section where it
/// stands; it never trades, as the merge pass family's own guard keeps it.
#[test]
fn a_paragraph_carrying_a_sectpr_keeps_its_place() {
    for mark in ["ins", "del"] {
        let other = ["del", "ins"].into_iter().find(|&k| k != mark).unwrap();
        let sect_mark = format!(
            r#"<w:p><w:pPr><w:sectPr><w:pgSz w:w="1" w:h="1"/></w:sectPr><w:rPr><w:{mark} w:id="1" {BY}/></w:rPr></w:pPr>{}</w:p>"#,
            wrapper(mark, 2, "break")
        );
        let (mut dom, root) = document(&format!("{sect_mark}{}", whole(other, 3, "mate")));
        let before = dom.serialize_element(root);
        reorder_deleted_paragraphs_before_inserted(&mut dom, root);
        assert_eq!(dom.serialize_element(root), before, "{mark}");
        assert_eq!(texts(&dom, body(&dom, root)), ["break", "mate"], "{mark}");
    }
}

/// The body's last paragraph, when it holds no run, is the story's final
/// pilcrow — which Word never deletes and the story-tail passes have
/// already settled. It keeps its place; an empty paragraph elsewhere still
/// counts by its mark.
#[test]
fn the_bodys_final_pilcrow_keeps_its_place() {
    let empty_deleted = marked("del", "");
    let (mut dom, root) = document(&format!("{}{empty_deleted}", whole("ins", 2, "new tail")));
    let (accepted, rejected) = (view(&dom, root, true), view(&dom, root, false));
    reorder_deleted_paragraphs_before_inserted(&mut dom, root);
    assert_eq!(texts(&dom, body(&dom, root)), ["new tail", ""]);
    assert_eq!(
        (view(&dom, root, true), view(&dom, root, false)),
        (accepted, rejected)
    );
}

/// A `w:tbl` ends a run: whole-revised paragraphs on either side of a
/// table never trade across it, and each side orders on its own.
#[test]
fn paragraphs_never_trade_across_a_table() {
    // The local names of the body's element children, and the text of its
    // direct paragraphs.
    let shape =
        |dom: &Dom, root: NodeId| (children(dom, body(dom, root)), texts(dom, body(dom, root)));
    let table = "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
    let (mut dom, root) = document(&format!(
        "{}{table}{}",
        whole("ins", 1, "before"),
        whole("del", 2, "after"),
    ));
    let (accepted, rejected) = (view(&dom, root, true), view(&dom, root, false));
    reorder_deleted_paragraphs_before_inserted(&mut dom, root);
    assert_eq!(
        shape(&dom, root),
        (
            vec![
                "p".to_string(),
                "tbl".to_string(),
                "p".to_string(),
                "sectPr".to_string(),
            ],
            vec!["before".to_string(), "after".to_string()]
        )
    );
    assert_eq!(
        (view(&dom, root, true), view(&dom, root, false)),
        (accepted, rejected)
    );

    // Each side of the table still orders on its own.
    let (mut dom, root) = document(&format!(
        "{}{}{table}{}{}",
        whole("ins", 1, "i1"),
        whole("del", 2, "d1"),
        whole("ins", 3, "i2"),
        whole("del", 4, "d2"),
    ));
    reorder_deleted_paragraphs_before_inserted(&mut dom, root);
    assert_eq!(
        shape(&dom, root),
        (
            vec![
                "p".to_string(),
                "p".to_string(),
                "tbl".to_string(),
                "p".to_string(),
                "p".to_string(),
                "sectPr".to_string()
            ],
            vec![
                "d1".to_string(),
                "i1".to_string(),
                "d2".to_string(),
                "i2".to_string()
            ]
        )
    );
}
