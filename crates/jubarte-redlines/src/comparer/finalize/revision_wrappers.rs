// SPDX-FileCopyrightText: 2026 Jandira Technologies, LLC
//
// SPDX-License-Identifier: Apache-2.0

//! Revision wrappers: insertion-before-deletion order, adjacent-revision coalescing, move simplification.

use super::*;

/// Word emits the INSERTION (new text) before the DELETION (old text) at a
/// replacement site; our LCS produces delete-then-insert. Swap an adjacent
/// same-author/date `w:del` immediately followed by `w:ins` into `w:ins` then
/// `w:del`, matching Word's order. Text-preserving: each of the delText / ins-text
/// streams keeps its own order (only their interleaving changes). Recurses.
///
/// A wrapper holding part of a field stays put: swapping a deleted field's
/// `end` with an inserted field's `begin` crosses the two fields, and Word
/// crashed opening the redline (English pair 57f96361×3832d290). A wrapper
/// whose fields are complete swaps whole: Word's redline of super_editor
/// hyperlink_multiple_runs × diff_before10 (corpus a098bd5ad9) writes the
/// inserted "One two" before the deleted HYPERLINK field.
/// Match Word's replacement order inside `w:r` and `w:p`: insertions
/// (`w:ins`) come before deletions (`w:del`).
///
/// When the LCS pairs a fully-deleted run with a fully-inserted
/// replacement run, the rebuilder may keep `w:del` ahead of `w:ins`
/// at a child position. Word always writes `ins+del`. The pass
/// reorders only at positions where two immediate siblings are one
/// each; deeper reordering is left for the multi-paragraph peel
/// ladder.
pub fn reorder_replacements_ins_before_del(dom: &mut Dom, node: NodeId) {
    let fld_char = W::fld_char();
    let mut i = 0usize;
    loop {
        let kids = dom.nodes(node);
        if i + 1 >= kids.len() {
            break;
        }
        let (a, b) = (kids[i], kids[i + 1]);
        let is_replacement = matches!(
            (revision_wrapper(dom, a), revision_wrapper(dom, b)),
            (Some((false, author_a, date_a)), Some((true, author_b, date_b)))
                if author_a == author_b && date_a == date_b
        ) && fields_complete(dom, a, &fld_char)
            && fields_complete(dom, b, &fld_char);
        if is_replacement {
            dom.remove(b);
            dom.add_before_self(a, b); // ins (b) now precedes del (a)
            i += 2;
        } else {
            i += 1;
        }
    }
    for c in dom.nodes(node) {
        if dom.is_element(c) {
            reorder_replacements_ins_before_del(dom, c);
        }
    }
}

/// Every field char under `node` belongs to a field that begins and ends
/// under it: no `fldChar` at all, or begins and ends that nest and balance.
#[doc(hidden)]
pub fn fields_complete(dom: &Dom, node: NodeId, fld_char: &XName) -> bool {
    let mut depth = 0i32;
    for fc in dom.descendants(node, Some(fld_char)) {
        match dom.attribute(fc, &W::name("fldCharType")) {
            Some("begin") => depth += 1,
            Some("end") => {
                depth -= 1;
                if depth < 0 {
                    return false;
                }
            }
            _ => {}
        }
    }
    depth == 0
}

/// A node that is one revision for the replacement swap: a `w:ins`
/// (`true`) or `w:del` (`false`) wrapper, or a `w:hyperlink` whose element
/// children are all one of them by one author and date (Word's redline of
/// super_editor hyperlink_multiple_runs × diff_before10, corpus a098bd5ad9,
/// writes "One two" inserted before the deleted "Click here now" link;
/// the swap saw a hyperlink, not a deletion, and left the link first).
#[doc(hidden)]
pub fn revision_wrapper(dom: &Dom, node: NodeId) -> Option<(bool, String, String)> {
    if !dom.is_element(node) {
        return None;
    }
    let own = |n: NodeId| {
        let kind = match dom.name(n).as_ref() {
            Some(name) if *name == W::ins() => true,
            Some(name) if *name == W::del() => false,
            _ => return None,
        };
        Some((
            kind,
            dom.attribute(n, &W::author()).unwrap_or("").to_string(),
            dom.attribute(n, &W::date()).unwrap_or("").to_string(),
        ))
    };
    if let Some(rev) = own(node) {
        return Some(rev);
    }
    if dom.name(node).as_ref() != Some(&W::hyperlink()) {
        return None;
    }
    let mut found: Option<(bool, String, String)> = None;
    for child in dom.nodes(node) {
        if !dom.is_element(child) {
            continue;
        }
        let rev = own(child)?;
        match &found {
            None => found = Some(rev),
            Some(first) if *first == rev => {}
            Some(_) => return None,
        }
    }
    found
}

/// Windows Word 2024 writes every replacement DELETION first and the
/// INSERTION after it, the reverse of Mac Word's order, which
/// [`reorder_replacements_ins_before_del`] writes. Under every element
/// except `w:rPr`, `w:pPr` and `w:trPr` (whose `w:ins` / `w:del` are
/// property revisions in schema order), take each maximal run of adjacent
/// revision children — a `w:ins` or `w:del`, or a `w:hyperlink` whose
/// element children are all one of them by one author and date, as
/// [`revision_wrapper`] classifies — and stable-partition each same
/// `(author, date)` group of it: every `w:del` before every `w:ins`, the
/// order inside each kind kept (#2555). One author's replacement never
/// fuses with another's: `[del A][ins B][del C][ins D]` with A/B by one
/// author and C/D by another stays two replacements, and a del and an ins
/// by different authors never trade, as the Mac pass only swaps
/// same-author/date pairs. The deleted and the inserted text each read as
/// before; only their interleaving changes. A group in which any wrapper
/// holds part of a field stays as it is ([`fields_complete`]): trading a
/// deleted field's `end` for an inserted field's `begin` crosses the two
/// fields, and Word refused the redline. Recurses into every child.
/// Idempotent and text-preserving (#2550).
pub fn reorder_deletions_before_insertions(dom: &mut Dom, node: NodeId) {
    let fld_char = W::fld_char();
    let property_bag = [W::r_pr(), W::p_pr(), W::tr_pr()];
    if dom.child_count(node) > 1 && dom.name(node).is_none_or(|n| !property_bag.contains(&n)) {
        let kids = dom.nodes(node);
        let mut start = 0;
        while start < kids.len() {
            if revision_wrapper(dom, kids[start]).is_none() {
                start += 1;
                continue;
            }
            let mut end = start;
            while end < kids.len() && revision_wrapper(dom, kids[end]).is_some() {
                end += 1;
            }
            let run = &kids[start..end];
            let mut group_start = 0;
            while group_start < run.len() {
                let (_, author, date) = revision_wrapper(dom, run[group_start]).unwrap();
                let mut group_end = group_start;
                while group_end < run.len()
                    && revision_wrapper(dom, run[group_end])
                        .is_some_and(|(_, a, d)| a == author && d == date)
                {
                    group_end += 1;
                }
                let group = &run[group_start..group_end];
                let kind =
                    |&c: &NodeId| revision_wrapper(dom, c).is_some_and(|(inserted, _, _)| inserted);
                let has_del = group.iter().any(|&c| !kind(&c));
                let has_ins = group.iter().any(&kind);
                if has_del && has_ins && group.iter().all(|&c| fields_complete(dom, c, &fld_char)) {
                    let (dels, inss): (Vec<NodeId>, Vec<NodeId>) =
                        group.iter().copied().partition(|&c| !kind(&c));
                    let order: Vec<NodeId> = dels.into_iter().chain(inss).collect();
                    let after = run
                        .get(group_end)
                        .copied()
                        .or_else(|| kids.get(end).copied());
                    order_siblings(dom, node, group, &order, after);
                }
                group_start = group_end;
            }
            start = end;
        }
    }
    for c in dom.nodes(node) {
        if dom.is_element(c) {
            reorder_deletions_before_insertions(dom, c);
        }
    }
}

/// Windows Word 2024 writes the DELETED paragraphs of a replaced block
/// before the inserted ones. [`reorder_replaced_blocks`] and the merge
/// passes write Mac Word's order, the inserted block first; the Windows
/// pipeline runs this pass in their place, once they have settled the
/// paragraphs. Under every `w:body`, `w:tc`, `w:footnote`, `w:endnote`,
/// `w:hdr`, `w:ftr`, `w:txbxContent`, `w:sdtContent` and `w:comment`, take
/// each maximal run of adjacent whole-deleted and whole-inserted
/// paragraphs that holds both kinds and stable-partition it, deletions
/// first (#2555: the merge pass [`merge_replaced_paragraphs_in`] already
/// walks the text-box, content-control and comment bodies, and so does
/// this pass; a `w:tbl` or `w:sdt` block between two paragraphs ends the
/// run, so no paragraph ever trades across one).
///
/// A paragraph is whole-deleted when its mark is deleted (`w:pPr/w:rPr/w:del`
/// and no `w:ins` there) and every `w:r` in it is covered by a `w:del`
/// between it and the paragraph (directly under the wrapper, or behind a
/// `w:hyperlink`, `w:sdt`, `w:smartTag` or `w:fldSimple` the wrapper holds,
/// as a text box's story is behind its drawing's run); whole-inserted is
/// the same with the two swapped (#2555). A paragraph with no run counts
/// by its mark alone. A paragraph whose `w:pPr` carries a `w:sectPr` — a
/// section break — counts as neither: it ends a section where it stands
/// and keeps its place, as the merge pass family's own guard keeps it
/// (`merge_replaced_in_container`). The body's last paragraph, when it
/// holds no run — the story's final pilcrow, which Word never deletes —
/// also keeps its place: the story-tail passes have already settled which
/// paragraph closes the story. Any other node, a mixed paragraph or one
/// whose mark and runs disagree ends the run. Stable, so accepting and
/// rejecting each read the same text as before. Idempotent (#2550).
pub fn reorder_deleted_paragraphs_before_inserted(dom: &mut Dom, root: NodeId) {
    let containers = [
        W::body(),
        W::tc(),
        W::footnote(),
        W::endnote(),
        W::hdr(),
        W::ftr(),
        W::txbx_content(),
        W::sdt_content(),
        W::comment(),
    ];
    for container in dom.descendants_and_self(root, None) {
        if !dom.name(container).is_some_and(|n| containers.contains(&n)) {
            continue;
        }
        let mut kids: Vec<(NodeId, Option<bool>)> = dom
            .nodes(container)
            .into_iter()
            .map(|k| (k, whole_revision_kind(dom, k)))
            .collect();
        if dom.name_is(container, &W::body())
            && let Some(last) = kids.iter().rposition(|&(k, _)| dom.name_is(k, &W::p()))
            && kids[last].1.is_some()
            && dom.descendants(kids[last].0, Some(&W::r())).is_empty()
        {
            kids[last].1 = None;
        }
        let mut start = 0;
        while start < kids.len() {
            if kids[start].1.is_none() {
                start += 1;
                continue;
            }
            let mut end = start;
            while end < kids.len() && kids[end].1.is_some() {
                end += 1;
            }
            let run = &kids[start..end];
            let deleted = run.iter().any(|&(_, k)| k == Some(false));
            let inserted = run.iter().any(|&(_, k)| k == Some(true));
            if deleted && inserted {
                let members: Vec<NodeId> = run.iter().map(|&(k, _)| k).collect();
                let order: Vec<NodeId> = run
                    .iter()
                    .filter(|&&(_, k)| k == Some(false))
                    .chain(run.iter().filter(|&&(_, k)| k == Some(true)))
                    .map(|&(k, _)| k)
                    .collect();
                order_siblings(
                    dom,
                    container,
                    &members,
                    &order,
                    kids.get(end).map(|&(k, _)| k),
                );
            }
            start = end;
        }
    }
}

/// `Some(false)` for a paragraph that is wholly deleted, `Some(true)` for
/// one wholly inserted and `None` for anything else; the paragraphs
/// [`reorder_deleted_paragraphs_before_inserted`] trades.
fn whole_revision_kind(dom: &Dom, p: NodeId) -> Option<bool> {
    if !dom.name_is(p, &W::p()) {
        return None;
    }
    let properties = dom.element(p, &W::p_pr());
    // A paragraph-level section break ends a section exactly where it
    // stands; trading it re-sections the document, so it counts as neither.
    // The merge pass family keeps the same guard (merge_replaced_in_container).
    if properties.is_some_and(|ppr| dom.element(ppr, &W::sect_pr()).is_some()) {
        return None;
    }
    let mark = properties.and_then(|ppr| dom.element(ppr, &W::r_pr()));
    let mark_has = |rev: &XName| mark.is_some_and(|m| dom.element(m, rev).is_some());
    let (mark_ins, mark_del) = (mark_has(&W::ins()), mark_has(&W::del()));
    if mark_ins == mark_del {
        return None;
    }
    let runs = dom.descendants(p, Some(&W::r()));
    let (wrapper, other) = if mark_ins {
        (W::ins(), W::del())
    } else {
        (W::del(), W::ins())
    };
    runs.iter()
        .all(|&r| run_covered_by(dom, r, p, &wrapper, &other))
        .then_some(mark_ins)
}

/// The first `w:ins` / `w:del` between the run `r` and its paragraph `p` is
/// `wrapper`: the run is as revised as one directly under the wrapper when
/// the wrapper holds it behind a `w:hyperlink`, `w:sdt`, `w:smartTag` or
/// `w:fldSimple` — or behind the drawing run of a text box whose story it
/// carries (#2555). A run that reaches `p`, or an opposite-kind wrapper,
/// first is not covered.
fn run_covered_by(dom: &Dom, r: NodeId, p: NodeId, wrapper: &XName, other: &XName) -> bool {
    let mut cur = dom.parent(r);
    while let Some(node) = cur {
        if node == p {
            return false;
        }
        if dom.name_is(node, other) {
            return false;
        }
        if dom.name_is(node, wrapper) {
            return true;
        }
        cur = dom.parent(node);
    }
    false
}

/// Move `order`, a permutation of the adjacent children `run` of `parent`,
/// into that order where the run stands. `after` is the sibling that follows
/// the run, when the run is not the parent's last children.
fn order_siblings(
    dom: &mut Dom,
    parent: NodeId,
    run: &[NodeId],
    order: &[NodeId],
    after: Option<NodeId>,
) {
    if order == run {
        return;
    }
    for &c in run {
        dom.remove(c);
    }
    for &c in order {
        match after {
            Some(next) => dom.add_before_self(next, c),
            None => dom.add(parent, c),
        }
    }
}

/// Merge consecutive sibling `w:ins` (and consecutive `w:del`) wrappers that
/// share the same `w:author` and `w:date` into a single wrapper, moving the
/// later wrapper's child runs into the first and removing the emptied wrapper.
/// Runs are kept intact (per-run formatting preserved); only the wrappers merge.
/// Word never emits adjacent same-status revision wrappers; this matches that
/// (PowerTools/ours otherwise inflate the w:ins/w:del element count ~2x).
/// Recurses through the whole tree. Text-preserving (the redlined text is unchanged).
/// Merge adjacent same-status `w:ins` / `w:del` wrappers under
/// `node` into one.
///
/// Word never emits two adjacent same-status revisions; the
/// redline pipeline can otherwise inflate the element count by ~2x.
/// Text-preserving — golden text-parity is unaffected. Walks the
/// children of `node` once, so a deep tree needs
/// [`coalesce_adjacent_runs`] or similar to repeat at lower levels.
pub fn coalesce_adjacent_revisions(dom: &mut Dom, node: NodeId) {
    let ins = W::ins();
    let del = W::del();
    let kids = dom.nodes(node);
    let mut prev: Option<NodeId> = None;
    for c in kids {
        let is_rev =
            dom.is_element(c) && matches!(dom.name(c), Some(ref n) if *n == ins || *n == del);
        if is_rev {
            if let Some(p) = prev {
                let same = dom.name(p) == dom.name(c)
                    && dom.attribute(p, &W::author()).map(|s| s.to_string())
                        == dom.attribute(c, &W::author()).map(|s| s.to_string())
                    && dom.attribute(p, &W::date()).map(|s| s.to_string())
                        == dom.attribute(c, &W::date()).map(|s| s.to_string());
                if same {
                    for gc in dom.nodes(c) {
                        dom.remove(gc);
                        dom.add(p, gc);
                    }
                    dom.remove(c);
                    continue; // prev unchanged; it may absorb the next sibling too
                }
            }
            prev = Some(c);
        } else {
            prev = None; // any non-revision sibling breaks adjacency
        }
    }
    for c in dom.nodes(node) {
        if dom.is_element(c) {
            coalesce_adjacent_revisions(dom, c);
        }
    }
}

/// M4.F — `SimplifyMoveMarkupToDelIns` (:3309), gated by settings.simplify_move_markup
/// (default off). Converts native move markup to del/ins for Word compatibility
/// (Issue #96): w:moveFrom→w:del, w:moveTo→w:ins (keeping author/date/id +
/// children), and removes all move range markers. Transform over `root`.
/// Convert native move markup (`w:moveFrom` / `w:moveTo`) into
/// `w:del` / `w:ins` (`SimplifyMoveMarkupToDelIns`,
/// WmlComparer.cs:2874-2878).
///
/// The `w:moveFrom` / `w:moveTo` half is rewritten to its
/// `W:author` / `W:date` / `W:id` (matching the C# `?.XAttribute`
/// pattern); stray attributes such as `rsid*`, `pt:Unid`, and any
/// future vendor attrs do not leak through. Gated on
/// [`WmlComparerSettings::simplify_move_markup`] (issue #96
/// workaround); default off.
pub fn simplify_move_markup_to_del_ins(dom: &mut Dom, root: NodeId) -> NodeId {
    simplify_move_transform(dom, root)
}

pub(super) fn simplify_move_transform(dom: &mut Dom, node: NodeId) -> NodeId {
    let name = dom.name(node).unwrap();
    let ranges = [
        W::move_from_range_start(),
        W::move_from_range_end(),
        W::move_to_range_start(),
        W::move_to_range_end(),
    ];
    let new_name = if name == W::move_from() {
        W::del()
    } else if name == W::move_to() {
        W::ins()
    } else {
        name.clone()
    };
    let ne = dom.new_element(new_name);
    // Oracle (WmlComparer.cs:2874-2878) propagates only w:author, w:date, w:id
    // — but only when rewriting w:moveFrom/w:moveTo. Recreated descendant
    // elements (w:r, w:p, …) must keep all of their attributes.
    if name == W::move_from() || name == W::move_to() {
        copy_move_id_attrs(dom, node, ne);
    } else {
        copy_attrs(dom, node, ne, false);
    }
    for c in dom.nodes(node) {
        if dom.is_element(c) {
            let cn = dom.name(c).unwrap();
            if ranges.contains(&cn) {
                continue; // drop range markers
            }
            let t = simplify_move_transform(dom, c);
            dom.add(ne, t);
        } else {
            let cc = dom.clone_subtree(c);
            dom.add(ne, cc);
        }
    }
    ne
}
