use super::model::{Document, Paragraph, Run};

/// Renders a Document to an HTML string for browser display.
/// Each paragraph gets `data-para="N"` and each run gets `data-run="M"`
/// to enable DOM <-> model position mapping.
pub fn render_to_html(doc: &Document) -> String {
    let mut html = String::with_capacity(4096);
    html.push_str("<div class=\"sofdocs-document\">");

    render_paragraphs(&mut html, &doc.body.paragraphs, doc);

    for table in &doc.body.tables {
        html.push_str("<table style=\"border-collapse:collapse;width:100%;\">");
        for row in &table.rows {
            html.push_str("<tr>");
            for cell in &row.cells {
                html.push_str("<td style=\"border:1px solid #ccc;padding:4px 8px;\">");
                for paragraph in &cell.paragraphs {
                    html.push_str("<p>");
                    for run in &paragraph.runs {
                        html.push_str(&html_escape(&run.text));
                    }
                    html.push_str("</p>");
                }
                html.push_str("</td>");
            }
            html.push_str("</tr>");
        }
        html.push_str("</table>");
    }

    // Shapes stubs
    for shape in &doc.shapes {
        html.push_str(&format!(
            "<div class=\"shape-placeholder\" style=\"border:2px dashed #999;padding:12px;margin:8px 0;text-align:center;color:#666;font-size:11px;border-radius:4px;width:{}px;min-height:{}px;display:flex;align-items:center;justify-content:center;\"><span>Shape: {} — {}</span></div>",
            shape.width_emu / 9525,
            shape.height_emu / 9525,
            html_escape(&shape.shape_type),
            if shape.text.is_empty() { "(no text)" } else { &shape.text }
        ));
    }

    // Footnotes stubs
    if !doc.footnotes.is_empty() {
        html.push_str("<hr style=\"margin-top:24px;border:none;border-top:1px solid #ccc;\"/>");
        html.push_str("<div class=\"footnotes\" style=\"font-size:9pt;color:#666;\">");
        for note in &doc.footnotes {
            html.push_str(&format!("<div class=\"footnote\" id=\"fn-{}\"><sup>{}</sup> ", note.id, note.id));
            for para in &note.paragraphs {
                for run in &para.runs {
                    html.push_str(&html_escape(&run.text));
                }
            }
            html.push_str("</div>");
        }
        html.push_str("</div>");
    }

    html.push_str("</div>");
    html
}

fn render_paragraphs(html: &mut String, paragraphs: &[Paragraph], doc: &Document) {
    let mut list_open: Option<String> = None;

    for (pi, paragraph) in paragraphs.iter().enumerate() {
        let is_list = paragraph.properties.numbering.is_some();

        if is_list {
            let list_tag = determine_list_tag(paragraph, doc);
            if list_open.as_deref() != Some(&list_tag) {
                if list_open.is_some() {
                    html.push_str(&format!("</{}>", list_open.as_deref().unwrap()));
                }
                html.push_str(&format!("<{list_tag}>"));
                list_open = Some(list_tag);
            }
            html.push_str(&format!("<li data-para=\"{pi}\">"));
            render_runs(html, &paragraph.runs, doc);
            html.push_str("</li>");
        } else {
            if let Some(ref tag) = list_open.take() {
                html.push_str(&format!("</{tag}>"));
            }

            if paragraph.properties.page_break_before {
                html.push_str("<div class=\"page-break\" style=\"page-break-before:always;border-top:1px dashed #ccc;margin:16px 0;\"></div>");
            }

            let tag = if paragraph.properties.heading_level > 0 {
                let level = paragraph.properties.heading_level.min(6);
                format!("h{level}")
            } else {
                "p".to_string()
            };

            html.push('<');
            html.push_str(&tag);
            html.push_str(&format!(" data-para=\"{pi}\""));

            let mut style = String::new();
            if let Some(ref align) = paragraph.properties.alignment {
                let css_align = match align.as_str() {
                    "center" => "center",
                    "right" | "end" => "right",
                    "both" | "justify" => "justify",
                    _ => "left",
                };
                style.push_str(&format!("text-align:{css_align};"));
            }
            if !style.is_empty() {
                html.push_str(&format!(" style=\"{style}\""));
            }
            html.push('>');

            if paragraph.runs.is_empty() {
                html.push_str("<span data-run=\"0\"><br></span>");
            } else {
                render_runs(html, &paragraph.runs, doc);
            }

            html.push_str("</");
            html.push_str(&tag);
            html.push('>');
        }
    }

    if let Some(ref tag) = list_open {
        html.push_str(&format!("</{tag}>"));
    }
}

fn determine_list_tag(para: &Paragraph, doc: &Document) -> String {
    if let Some(ref num) = para.properties.numbering {
        for def in &doc.numbering_definitions {
            for lvl in &def.levels {
                if lvl.level == num.level {
                    if lvl.num_fmt == "bullet" {
                        return "ul".to_string();
                    } else {
                        return "ol".to_string();
                    }
                }
            }
        }
    }
    "ul".to_string()
}

fn render_runs(html: &mut String, runs: &[Run], doc: &Document) {
    for (ri, run) in runs.iter().enumerate() {
        if let Some(ref img) = run.image {
            let (data, content_type) = find_image_info(doc, &img.r_id)
                .unwrap_or((&img.data, &img.content_type));
            if !data.is_empty() {
                let b64 = base64_encode(data);
                let w_px = img.width_emu / 9525;
                let h_px = img.height_emu / 9525;
                let ct = if content_type.is_empty() { "image/png" } else { content_type };
                html.push_str(&format!(
                    "<img data-run=\"{ri}\" src=\"data:{};base64,{}\" style=\"width:{}px;height:{}px;\" alt=\"{}\"/>",
                    ct,
                    b64,
                    w_px,
                    h_px,
                    html_escape(img.description.as_deref().unwrap_or(""))
                ));
            }
            continue;
        }

        let is_link = run.hyperlink.is_some();
        if is_link {
            let link = run.hyperlink.as_ref().unwrap();
            html.push_str(&format!(
                "<a href=\"{}\" data-run=\"{ri}\" target=\"_blank\" rel=\"noopener\" style=\"color:#0563C1;text-decoration:underline;\">",
                html_escape(&link.url)
            ));
        }

        let mut span_style = String::new();
        if run.style.bold {
            span_style.push_str("font-weight:bold;");
        }
        if run.style.italic {
            span_style.push_str("font-style:italic;");
        }
        if run.style.underline && run.style.strikethrough {
            span_style.push_str("text-decoration:underline line-through;");
        } else if run.style.underline {
            span_style.push_str("text-decoration:underline;");
        } else if run.style.strikethrough {
            span_style.push_str("text-decoration:line-through;");
        }
        if let Some(ref font) = run.style.font_family {
            span_style.push_str(&format!("font-family:'{font}',sans-serif;"));
        }
        if let Some(size) = run.style.font_size_pt {
            span_style.push_str(&format!("font-size:{size}pt;"));
        }
        if let Some(ref color) = run.style.color {
            span_style.push_str(&format!("color:#{color};"));
        }
        if let Some(ref highlight) = run.style.highlight {
            let css_color = highlight_to_css(highlight);
            span_style.push_str(&format!("background-color:{css_color};"));
        }

        let use_sub_sup = run.style.vertical_align.is_some();
        let sub_sup_tag = match &run.style.vertical_align {
            Some(super::model::VerticalAlign::Superscript) => "sup",
            Some(super::model::VerticalAlign::Subscript) => "sub",
            None => "",
        };

        if use_sub_sup {
            html.push_str(&format!("<{sub_sup_tag}>"));
        }

        if !span_style.is_empty() {
            html.push_str(&format!("<span data-run=\"{ri}\" style=\"{span_style}\">"));
        } else if !is_link {
            html.push_str(&format!("<span data-run=\"{ri}\">"));
        }

        if run.text.is_empty() {
            html.push_str("<br>");
        } else {
            html.push_str(&html_escape(&run.text));
        }

        if !span_style.is_empty() || !is_link {
            html.push_str("</span>");
        }

        if use_sub_sup {
            html.push_str(&format!("</{sub_sup_tag}>"));
        }
        if is_link {
            html.push_str("</a>");
        }
    }
}

fn highlight_to_css(ooxml_color: &str) -> &str {
    match ooxml_color {
        "yellow" => "#FFFF00",
        "green" => "#00FF00",
        "cyan" => "#00FFFF",
        "magenta" => "#FF00FF",
        "blue" => "#0000FF",
        "red" => "#FF0000",
        "darkBlue" => "#00008B",
        "darkCyan" => "#008B8B",
        "darkGreen" => "#006400",
        "darkMagenta" => "#8B008B",
        "darkRed" => "#8B0000",
        "darkYellow" => "#808000",
        "darkGray" => "#A9A9A9",
        "lightGray" => "#D3D3D3",
        "black" => "#000000",
        "white" => "#FFFFFF",
        _ => ooxml_color,
    }
}

fn find_image_info<'a>(doc: &'a Document, r_id: &str) -> Option<(&'a [u8], &'a str)> {
    doc.images
        .iter()
        .find(|img| img.r_id == r_id)
        .map(|img| (img.data.as_slice(), img.content_type.as_str()))
}

fn base64_encode(data: &[u8]) -> String {
    const CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut result = String::with_capacity((data.len() + 2) / 3 * 4);
    for chunk in data.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = if chunk.len() > 1 { chunk[1] as u32 } else { 0 };
        let b2 = if chunk.len() > 2 { chunk[2] as u32 } else { 0 };
        let triple = (b0 << 16) | (b1 << 8) | b2;
        result.push(CHARS[((triple >> 18) & 0x3F) as usize] as char);
        result.push(CHARS[((triple >> 12) & 0x3F) as usize] as char);
        if chunk.len() > 1 {
            result.push(CHARS[((triple >> 6) & 0x3F) as usize] as char);
        } else {
            result.push('=');
        }
        if chunk.len() > 2 {
            result.push(CHARS[(triple & 0x3F) as usize] as char);
        } else {
            result.push('=');
        }
    }
    result
}

fn html_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}
