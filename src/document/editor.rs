use super::model::*;

/// Represents a position in the document: paragraph index + character offset within that paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct DocPosition {
    pub paragraph: usize,
    pub offset: usize,
}

/// A selection range within the document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, serde::Serialize, serde::Deserialize)]
pub struct DocSelection {
    pub start: DocPosition,
    pub end: DocPosition,
}

impl DocSelection {
    pub fn is_collapsed(&self) -> bool {
        self.start == self.end
    }

    /// Ensure start <= end.
    pub fn normalized(&self) -> Self {
        if self.start.paragraph > self.end.paragraph
            || (self.start.paragraph == self.end.paragraph && self.start.offset > self.end.offset)
        {
            Self {
                start: self.end,
                end: self.start,
            }
        } else {
            *self
        }
    }
}

/// An undoable edit operation.
#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub enum EditOp {
    InsertText {
        position: DocPosition,
        text: String,
    },
    DeleteRange {
        selection: DocSelection,
        deleted_content: Vec<Paragraph>,
    },
    ApplyStyle {
        selection: DocSelection,
        style_change: StyleChange,
        previous_runs: Vec<(usize, Vec<Run>)>,
    },
    SetAlignment {
        paragraph: usize,
        new_alignment: Option<String>,
        old_alignment: Option<String>,
    },
    SplitParagraph {
        position: DocPosition,
    },
}

#[derive(Debug, Clone, serde::Serialize, serde::Deserialize)]
pub enum StyleChange {
    ToggleBold,
    ToggleItalic,
    ToggleUnderline,
    ToggleStrikethrough,
    SetFontFamily(String),
    SetFontSize(f32),
    SetColor(String),
}

/// Undo/redo stack holding edit operations.
pub struct UndoStack {
    undo: Vec<EditOp>,
    redo: Vec<EditOp>,
}

impl UndoStack {
    pub fn new() -> Self {
        Self {
            undo: Vec::new(),
            redo: Vec::new(),
        }
    }

    pub fn push(&mut self, op: EditOp) {
        self.undo.push(op);
        self.redo.clear();
    }

    pub fn can_undo(&self) -> bool {
        !self.undo.is_empty()
    }

    pub fn can_redo(&self) -> bool {
        !self.redo.is_empty()
    }

    pub fn pop_undo(&mut self) -> Option<EditOp> {
        self.undo.pop()
    }

    pub fn push_redo(&mut self, op: EditOp) {
        self.redo.push(op);
    }

    pub fn pop_redo(&mut self) -> Option<EditOp> {
        self.redo.pop()
    }
}

impl Default for UndoStack {
    fn default() -> Self {
        Self::new()
    }
}

/// Get the total character count in a paragraph (sum of all run texts).
fn paragraph_char_count(para: &Paragraph) -> usize {
    para.runs.iter().map(|r| r.text.len()).sum()
}

/// Find which run contains the given character offset, and the offset within that run.
fn find_run_at_offset(para: &Paragraph, offset: usize) -> (usize, usize) {
    let mut remaining = offset;
    for (i, run) in para.runs.iter().enumerate() {
        if remaining <= run.text.len() {
            return (i, remaining);
        }
        remaining -= run.text.len();
    }
    let last = para.runs.len().saturating_sub(1);
    let last_len = para.runs.last().map(|r| r.text.len()).unwrap_or(0);
    (last, last_len)
}

/// Insert text at a position. Returns the end position after insertion.
pub fn insert_text(doc: &mut Document, pos: DocPosition, text: &str) -> DocPosition {
    if pos.paragraph >= doc.body.paragraphs.len() {
        doc.body.paragraphs.push(Paragraph::default());
    }
    let para = &mut doc.body.paragraphs[pos.paragraph];

    if para.runs.is_empty() {
        para.runs.push(Run {
            text: text.to_string(),
            style: RunStyle::default(),
        });
        return DocPosition {
            paragraph: pos.paragraph,
            offset: text.len(),
        };
    }

    let (run_idx, run_offset) = find_run_at_offset(para, pos.offset);
    para.runs[run_idx].text.insert_str(run_offset, text);

    DocPosition {
        paragraph: pos.paragraph,
        offset: pos.offset + text.len(),
    }
}

/// Delete the content in a selection. Returns the collapsed position.
pub fn delete_range(doc: &mut Document, sel: DocSelection) -> (DocPosition, Vec<Paragraph>) {
    let sel = sel.normalized();

    if sel.is_collapsed() {
        return (sel.start, Vec::new());
    }

    if sel.start.paragraph == sel.end.paragraph {
        let para = &mut doc.body.paragraphs[sel.start.paragraph];
        let deleted = extract_paragraph_range(para, sel.start.offset, sel.end.offset);
        delete_chars_in_paragraph(para, sel.start.offset, sel.end.offset);
        return (sel.start, vec![deleted]);
    }

    let mut deleted_paras = Vec::new();

    // Save the deleted content from the first paragraph (from start.offset to end)
    let first_para = &doc.body.paragraphs[sel.start.paragraph];
    let first_len = paragraph_char_count(first_para);
    deleted_paras.push(extract_paragraph_range(
        first_para,
        sel.start.offset,
        first_len,
    ));

    // Save full middle paragraphs
    for i in (sel.start.paragraph + 1)..sel.end.paragraph {
        deleted_paras.push(doc.body.paragraphs[i].clone());
    }

    // Save deleted content from last paragraph
    let last_para = &doc.body.paragraphs[sel.end.paragraph];
    deleted_paras.push(extract_paragraph_range(last_para, 0, sel.end.offset));

    // Now perform the deletion: keep the tail of the last paragraph
    let last_para = &doc.body.paragraphs[sel.end.paragraph];
    let last_len = paragraph_char_count(last_para);
    let tail_runs = extract_runs_from_offset(last_para, sel.end.offset, last_len);

    // Truncate the first paragraph at start.offset
    delete_chars_in_paragraph(
        &mut doc.body.paragraphs[sel.start.paragraph],
        sel.start.offset,
        first_len,
    );

    // Append tail of last paragraph to first
    for run in tail_runs {
        if !run.text.is_empty() {
            doc.body.paragraphs[sel.start.paragraph].runs.push(run);
        }
    }

    // Remove the middle + last paragraphs
    let remove_count = sel.end.paragraph - sel.start.paragraph;
    for _ in 0..remove_count {
        doc.body.paragraphs.remove(sel.start.paragraph + 1);
    }

    (sel.start, deleted_paras)
}

/// Split a paragraph at a position (Enter key).
pub fn split_paragraph(doc: &mut Document, pos: DocPosition) {
    if pos.paragraph >= doc.body.paragraphs.len() {
        doc.body.paragraphs.push(Paragraph::default());
        return;
    }

    let para = &doc.body.paragraphs[pos.paragraph];
    let total_len = paragraph_char_count(para);
    let tail_runs = extract_runs_from_offset(para, pos.offset, total_len);

    // Truncate current paragraph
    delete_chars_in_paragraph(
        &mut doc.body.paragraphs[pos.paragraph],
        pos.offset,
        total_len,
    );

    // Create new paragraph with tail content
    let new_para = Paragraph {
        properties: doc.body.paragraphs[pos.paragraph].properties.clone(),
        runs: tail_runs,
    };

    doc.body.paragraphs.insert(pos.paragraph + 1, new_para);
}

/// Apply a style change to a selection.
pub fn apply_style(
    doc: &mut Document,
    sel: DocSelection,
    change: &StyleChange,
) -> Vec<(usize, Vec<Run>)> {
    let sel = sel.normalized();
    let mut previous_runs = Vec::new();

    let start_para = sel.start.paragraph;
    let end_para = sel.end.paragraph.min(doc.body.paragraphs.len().saturating_sub(1));

    for pi in start_para..=end_para {
        let para = &doc.body.paragraphs[pi];
        previous_runs.push((pi, para.runs.clone()));

        let para_len = paragraph_char_count(para);
        let start_offset = if pi == start_para { sel.start.offset } else { 0 };
        let end_offset = if pi == end_para {
            sel.end.offset
        } else {
            para_len
        };

        split_runs_at_boundaries(&mut doc.body.paragraphs[pi], start_offset, end_offset);

        let para = &mut doc.body.paragraphs[pi];
        let mut char_pos = 0;
        for run in &mut para.runs {
            let run_end = char_pos + run.text.len();
            if char_pos >= start_offset && run_end <= end_offset && !run.text.is_empty() {
                apply_style_change(&mut run.style, change);
            }
            char_pos = run_end;
        }
    }

    previous_runs
}

/// Set paragraph alignment.
pub fn set_alignment(
    doc: &mut Document,
    paragraph: usize,
    alignment: Option<String>,
) -> Option<String> {
    if paragraph >= doc.body.paragraphs.len() {
        return None;
    }
    let old = doc.body.paragraphs[paragraph].properties.alignment.clone();
    doc.body.paragraphs[paragraph].properties.alignment = alignment;
    old
}

fn apply_style_change(style: &mut RunStyle, change: &StyleChange) {
    match change {
        StyleChange::ToggleBold => style.bold = !style.bold,
        StyleChange::ToggleItalic => style.italic = !style.italic,
        StyleChange::ToggleUnderline => style.underline = !style.underline,
        StyleChange::ToggleStrikethrough => style.strikethrough = !style.strikethrough,
        StyleChange::SetFontFamily(f) => style.font_family = Some(f.clone()),
        StyleChange::SetFontSize(s) => style.font_size_pt = Some(*s),
        StyleChange::SetColor(c) => style.color = Some(c.clone()),
    }
}

/// Split runs at character boundaries so that the range [start..end] aligns with run boundaries.
fn split_runs_at_boundaries(para: &mut Paragraph, start: usize, end: usize) {
    split_run_at(para, end);
    split_run_at(para, start);
}

fn split_run_at(para: &mut Paragraph, offset: usize) {
    let mut char_pos = 0;
    for i in 0..para.runs.len() {
        let run_len = para.runs[i].text.len();
        if char_pos == offset || char_pos + run_len <= offset {
            char_pos += run_len;
            continue;
        }
        if offset > char_pos && offset < char_pos + run_len {
            let split_at = offset - char_pos;
            let tail_text = para.runs[i].text[split_at..].to_string();
            para.runs[i].text.truncate(split_at);
            let new_run = Run {
                text: tail_text,
                style: para.runs[i].style.clone(),
            };
            para.runs.insert(i + 1, new_run);
            return;
        }
        char_pos += run_len;
    }
}

fn delete_chars_in_paragraph(para: &mut Paragraph, start: usize, end: usize) {
    split_runs_at_boundaries(para, start, end);

    let mut char_pos = 0;
    para.runs.retain(|run| {
        let run_end = char_pos + run.text.len();
        let keep = run_end <= start || char_pos >= end;
        char_pos = run_end;
        keep
    });
}

fn extract_paragraph_range(para: &Paragraph, start: usize, end: usize) -> Paragraph {
    let runs = extract_runs_from_offset(para, start, end);
    Paragraph {
        properties: para.properties.clone(),
        runs,
    }
}

fn extract_runs_from_offset(para: &Paragraph, start: usize, end: usize) -> Vec<Run> {
    let mut result = Vec::new();
    let mut char_pos = 0;

    for run in &para.runs {
        let run_end = char_pos + run.text.len();

        if run_end <= start || char_pos >= end {
            char_pos = run_end;
            continue;
        }

        let slice_start = if char_pos < start {
            start - char_pos
        } else {
            0
        };
        let slice_end = if run_end > end {
            end - char_pos
        } else {
            run.text.len()
        };

        let text = run.text[slice_start..slice_end].to_string();
        if !text.is_empty() {
            result.push(Run {
                text,
                style: run.style.clone(),
            });
        }

        char_pos = run_end;
    }

    result
}

/// Perform an undo operation. Returns the inverse op for redo.
pub fn undo(doc: &mut Document, undo_stack: &mut UndoStack) -> bool {
    let op = match undo_stack.pop_undo() {
        Some(op) => op,
        None => return false,
    };

    let redo_op = execute_inverse(doc, &op);
    undo_stack.push_redo(redo_op);
    true
}

/// Perform a redo operation.
pub fn redo(doc: &mut Document, undo_stack: &mut UndoStack) -> bool {
    let op = match undo_stack.pop_redo() {
        Some(op) => op,
        None => return false,
    };

    let inverse = execute_inverse(doc, &op);
    undo_stack.undo.push(inverse);
    true
}

fn execute_inverse(doc: &mut Document, op: &EditOp) -> EditOp {
    match op {
        EditOp::InsertText { position, text } => {
            let end = DocPosition {
                paragraph: position.paragraph,
                offset: position.offset + text.len(),
            };
            let sel = DocSelection {
                start: *position,
                end,
            };
            let (_, deleted) = delete_range(doc, sel);
            EditOp::DeleteRange {
                selection: sel,
                deleted_content: deleted,
            }
        }
        EditOp::DeleteRange {
            selection,
            deleted_content,
        } => {
            let sel = selection.normalized();
            let mut restored_text = String::new();
            for (i, para) in deleted_content.iter().enumerate() {
                if i > 0 {
                    restored_text.push('\n');
                }
                for run in &para.runs {
                    restored_text.push_str(&run.text);
                }
            }
            insert_text(doc, sel.start, &restored_text);
            EditOp::InsertText {
                position: sel.start,
                text: restored_text,
            }
        }
        EditOp::ApplyStyle {
            selection,
            style_change,
            previous_runs,
        } => {
            let current_runs: Vec<(usize, Vec<Run>)> = previous_runs
                .iter()
                .map(|(pi, _)| {
                    let runs = doc.body.paragraphs[*pi].runs.clone();
                    (*pi, runs)
                })
                .collect();

            for (pi, runs) in previous_runs {
                if *pi < doc.body.paragraphs.len() {
                    doc.body.paragraphs[*pi].runs = runs.clone();
                }
            }

            EditOp::ApplyStyle {
                selection: *selection,
                style_change: style_change.clone(),
                previous_runs: current_runs,
            }
        }
        EditOp::SetAlignment {
            paragraph,
            new_alignment,
            old_alignment,
        } => {
            set_alignment(doc, *paragraph, old_alignment.clone());
            EditOp::SetAlignment {
                paragraph: *paragraph,
                new_alignment: old_alignment.clone(),
                old_alignment: new_alignment.clone(),
            }
        }
        EditOp::SplitParagraph { position } => {
            // Merge the two paragraphs back
            if position.paragraph + 1 < doc.body.paragraphs.len() {
                let next_runs = doc.body.paragraphs[position.paragraph + 1].runs.clone();
                doc.body.paragraphs.remove(position.paragraph + 1);
                doc.body.paragraphs[position.paragraph]
                    .runs
                    .extend(next_runs);
            }
            EditOp::SplitParagraph {
                position: *position,
            }
        }
    }
}
