use super::model::Document;

/// Renders a Document to an HTML string for browser display.
/// Each paragraph gets `data-para="N"` and each run gets `data-run="M"`
/// to enable DOM ↔ model position mapping.
pub fn render_to_html(doc: &Document) -> String {
    let mut html = String::with_capacity(4096);
    html.push_str("<div class=\"sofdocs-document\">");

    for (pi, paragraph) in doc.body.paragraphs.iter().enumerate() {
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
        }

        for (ri, run) in paragraph.runs.iter().enumerate() {
            let mut span_style = String::new();
            if run.style.bold {
                span_style.push_str("font-weight:bold;");
            }
            if run.style.italic {
                span_style.push_str("font-style:italic;");
            }
            if run.style.underline {
                span_style.push_str("text-decoration:underline;");
            }
            if run.style.strikethrough {
                if span_style.contains("text-decoration:") {
                    span_style = span_style.replace("text-decoration:underline;", "text-decoration:underline line-through;");
                } else {
                    span_style.push_str("text-decoration:line-through;");
                }
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

            if !span_style.is_empty() {
                html.push_str(&format!("<span data-run=\"{ri}\" style=\"{span_style}\">"));
            } else {
                html.push_str(&format!("<span data-run=\"{ri}\">"));
            }

            if run.text.is_empty() {
                html.push_str("<br>");
            } else {
                html.push_str(&html_escape(&run.text));
            }
            html.push_str("</span>");
        }

        html.push_str("</");
        html.push_str(&tag);
        html.push('>');
    }

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

    html.push_str("</div>");
    html
}

fn html_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}
