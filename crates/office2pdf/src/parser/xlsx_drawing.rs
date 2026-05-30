use std::collections::{HashMap, HashSet};
use std::io::Cursor;

use crate::ir::Chart;
use crate::ir::{FixedElement, FixedElementKind, ImageData, ImageFormat, Margins};
use crate::parser::chart::parse_chart_xml;
use crate::parser::xml_util;

use super::xlsx_cells::{DEFAULT_COLUMN_WIDTH, SheetContext, column_width_to_pt};

const EMU_PER_POINT: f64 = 12_700.0;
const DEFAULT_ROW_HEIGHT_PT: f64 = 15.0;

#[derive(Debug, Clone, Default)]
struct Marker {
    col: u32,
    col_off_emu: i64,
    row: u32,
    row_off_emu: i64,
}

#[derive(Debug, Clone)]
struct PictureAnchor {
    from: Marker,
    to: Option<Marker>,
    ext: Option<(i64, i64)>,
    rid: String,
}

fn emu_to_pt(value: i64) -> f64 {
    value as f64 / EMU_PER_POINT
}

fn image_format_from_path(path: &str) -> Option<ImageFormat> {
    let ext = path.rsplit('.').next()?.to_ascii_lowercase();
    match ext.as_str() {
        "png" => Some(ImageFormat::Png),
        "jpg" | "jpeg" => Some(ImageFormat::Jpeg),
        "gif" => Some(ImageFormat::Gif),
        "bmp" => Some(ImageFormat::Bmp),
        "tif" | "tiff" => Some(ImageFormat::Tiff),
        "svg" => Some(ImageFormat::Svg),
        _ => None,
    }
}

fn read_zip_entry_bytes(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    path: &str,
) -> Option<Vec<u8>> {
    let mut entry = archive.by_name(path).ok()?;
    let mut bytes = Vec::new();
    std::io::Read::read_to_end(&mut entry, &mut bytes).ok()?;
    Some(bytes)
}

fn row_height_pt(sheet: &umya_spreadsheet::Worksheet, row: u32) -> f64 {
    sheet
        .get_row_dimension(&row)
        .filter(|r| *r.get_custom_height())
        .map(|r| *r.get_height())
        .unwrap_or(DEFAULT_ROW_HEIGHT_PT)
}

fn column_width_pt(sheet: &umya_spreadsheet::Worksheet, col: u32) -> f64 {
    sheet
        .get_column_dimension_by_number(&col)
        .map(|c| column_width_to_pt(*c.get_width()))
        .unwrap_or_else(|| column_width_to_pt(DEFAULT_COLUMN_WIDTH))
}

fn x_from_marker(
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
    margins: &Margins,
    marker: &Marker,
) -> f64 {
    // Drawing anchors are zero-based; SheetContext coordinates are one-based.
    let col = marker.col + 1;
    let mut x = margins.left;
    for c in ctx.col_start..col {
        x += column_width_pt(sheet, c);
    }
    x + emu_to_pt(marker.col_off_emu)
}

fn y_from_marker(
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
    margins: &Margins,
    marker: &Marker,
) -> f64 {
    let row = marker.row + 1;
    let mut y = margins.top;
    for r in ctx.row_start..row {
        y += row_height_pt(sheet, r);
    }
    y + emu_to_pt(marker.row_off_emu)
}

fn anchor_size_pt(sheet: &umya_spreadsheet::Worksheet, anchor: &PictureAnchor) -> (f64, f64) {
    if let Some((cx, cy)) = anchor.ext {
        return (emu_to_pt(cx), emu_to_pt(cy));
    }

    let Some(to) = anchor.to.as_ref() else {
        return (0.0, 0.0);
    };

    let from_col = anchor.from.col + 1;
    let to_col = to.col + 1;
    let from_row = anchor.from.row + 1;
    let to_row = to.row + 1;

    let mut width = -emu_to_pt(anchor.from.col_off_emu) + emu_to_pt(to.col_off_emu);
    for c in from_col..to_col {
        width += column_width_pt(sheet, c);
    }

    let mut height = -emu_to_pt(anchor.from.row_off_emu) + emu_to_pt(to.row_off_emu);
    for r in from_row..to_row {
        height += row_height_pt(sheet, r);
    }

    // Keep obviously malformed anchors out of Typst.
    (width.max(0.0), height.max(0.0))
}

/// Extract worksheet pictures as fixed overlays.
///
/// This intentionally handles raster/vector pictures (`xdr:pic`) separately from
/// charts. Existing chart extraction renders charts as data tables; pictures
/// must stay as images because templates often use them for logos/stamps/TTD.
pub(super) fn extract_sheet_picture_overlays(
    data: &[u8],
    sheet_name: &str,
    sheet: &umya_spreadsheet::Worksheet,
    ctx: &SheetContext,
    margins: &Margins,
) -> Vec<FixedElement> {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return Vec::new();
    };

    let workbook_xml = read_zip_entry_string(&mut archive, "xl/workbook.xml");
    let sheet_rids = parse_workbook_sheet_rids(&workbook_xml);
    let Some((_, sheet_rid)) = sheet_rids.iter().find(|(name, _)| name == sheet_name) else {
        return Vec::new();
    };

    let workbook_rels_xml = read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels");
    let rid_to_target = parse_rels_targets(&workbook_rels_xml);
    let Some(sheet_target) = rid_to_target.get(sheet_rid) else {
        return Vec::new();
    };

    let sheet_full_path = format!("xl/{sheet_target}");
    let sheet_filename = sheet_full_path.rsplit('/').next().unwrap_or(sheet_target);
    let sheet_rels_path = format!("xl/worksheets/_rels/{sheet_filename}.rels");
    let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path);
    if sheet_rels_xml.is_empty() {
        return Vec::new();
    }

    let mut overlays = Vec::new();
    for drawing_target in parse_rels_by_type(&sheet_rels_xml, "drawing") {
        let drawing_path = resolve_relative_xl_path("xl/worksheets", &drawing_target);
        let drawing_xml = read_zip_entry_string(&mut archive, &drawing_path);
        if drawing_xml.is_empty() {
            continue;
        }

        let drawing_filename = drawing_path.rsplit('/').next().unwrap_or(&drawing_path);
        let drawing_dir = drawing_path
            .rsplit_once('/')
            .map(|(d, _)| d)
            .unwrap_or("xl/drawings");
        let drawing_rels_path = format!("{drawing_dir}/_rels/{drawing_filename}.rels");
        let drawing_rels_xml = read_zip_entry_string(&mut archive, &drawing_rels_path);
        let drawing_rid_targets = parse_rels_targets(&drawing_rels_xml);

        for anchor in parse_drawing_picture_anchors(&drawing_xml) {
            let Some(target) = drawing_rid_targets.get(&anchor.rid) else {
                continue;
            };
            let image_path = resolve_relative_xl_path(drawing_dir, target);
            let Some(format) = image_format_from_path(&image_path) else {
                continue;
            };
            let Some(bytes) = read_zip_entry_bytes(&mut archive, &image_path) else {
                continue;
            };

            let x = x_from_marker(sheet, ctx, margins, &anchor.from);
            let y = y_from_marker(sheet, ctx, margins, &anchor.from);
            let (width, height) = anchor_size_pt(sheet, &anchor);
            if width <= 0.0 || height <= 0.0 {
                continue;
            }

            overlays.push(FixedElement {
                x,
                y,
                width,
                height,
                kind: FixedElementKind::Image(ImageData {
                    data: bytes,
                    format,
                    width: Some(width),
                    height: Some(height),
                    crop: None,
                    stroke: None,
                }),
            });
        }
    }

    overlays
}

fn parse_drawing_picture_anchors(xml: &str) -> Vec<PictureAnchor> {
    let mut result = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    let mut current_anchor: Option<PictureAnchor> = None;
    let mut marker_target: Option<&'static str> = None;
    let mut current_marker: Option<Marker> = None;
    let mut current_field: Option<&'static str> = None;
    let mut in_pic = false;

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e)) => match e.local_name().as_ref() {
                b"twoCellAnchor" | b"oneCellAnchor" => {
                    current_anchor = Some(PictureAnchor {
                        from: Marker::default(),
                        to: None,
                        ext: None,
                        rid: String::new(),
                    });
                }
                b"from" if current_anchor.is_some() => {
                    marker_target = Some("from");
                    current_marker = Some(Marker::default());
                }
                b"to" if current_anchor.is_some() => {
                    marker_target = Some("to");
                    current_marker = Some(Marker::default());
                }
                b"col" => current_field = Some("col"),
                b"colOff" => current_field = Some("colOff"),
                b"row" => current_field = Some("row"),
                b"rowOff" => current_field = Some("rowOff"),
                b"ext" => {
                    if let Some(anchor) = current_anchor.as_mut() {
                        let cx = xml_util::get_attr_i64(e, b"cx").unwrap_or(0);
                        let cy = xml_util::get_attr_i64(e, b"cy").unwrap_or(0);
                        if cx > 0 && cy > 0 {
                            anchor.ext = Some((cx, cy));
                        }
                    }
                }
                b"pic" => in_pic = true,
                b"blip" if in_pic => {
                    if let Some(anchor) = current_anchor.as_mut()
                        && let Some(rid) = xml_util::get_attr_str(e, b"embed")
                    {
                        anchor.rid = rid;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"ext" => {
                    if let Some(anchor) = current_anchor.as_mut() {
                        let cx = xml_util::get_attr_i64(e, b"cx").unwrap_or(0);
                        let cy = xml_util::get_attr_i64(e, b"cy").unwrap_or(0);
                        if cx > 0 && cy > 0 {
                            anchor.ext = Some((cx, cy));
                        }
                    }
                }
                b"blip" if in_pic => {
                    if let Some(anchor) = current_anchor.as_mut()
                        && let Some(rid) = xml_util::get_attr_str(e, b"embed")
                    {
                        anchor.rid = rid;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Text(ref t)) => {
                if let (Some(marker), Some(field)) = (current_marker.as_mut(), current_field)
                    && let Ok(text) = t.xml_content()
                    && let Ok(value) = text.trim().parse::<i64>()
                {
                    match field {
                        "col" => marker.col = value.max(0) as u32,
                        "colOff" => marker.col_off_emu = value,
                        "row" => marker.row = value.max(0) as u32,
                        "rowOff" => marker.row_off_emu = value,
                        _ => {}
                    }
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => match e.local_name().as_ref() {
                b"from" | b"to" => {
                    if let (Some(anchor), Some(target), Some(marker)) = (
                        current_anchor.as_mut(),
                        marker_target.take(),
                        current_marker.take(),
                    ) {
                        if target == "from" {
                            anchor.from = marker;
                        } else {
                            anchor.to = Some(marker);
                        }
                    }
                    current_field = None;
                }
                b"col" | b"colOff" | b"row" | b"rowOff" => {
                    current_field = None;
                }
                b"pic" => in_pic = false,
                b"twoCellAnchor" | b"oneCellAnchor" => {
                    if let Some(anchor) = current_anchor.take()
                        && !anchor.rid.is_empty()
                    {
                        result.push(anchor);
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    result
}

/// Extract charts from the XLSX ZIP with their anchor positions per sheet.
///
/// Returns a map from sheet name → list of (anchor_row, Chart).
/// Charts with drawing anchors get positioned at their anchor row.
/// Charts without anchors (no drawing reference found) use `u32::MAX`
/// as a sentinel to place them at the end of the sheet.
pub(super) fn extract_charts_with_anchors(data: &[u8]) -> HashMap<String, Vec<(u32, Chart)>> {
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return HashMap::new();
    };

    // Step 1: Read workbook.xml to get sheet name → rId mapping
    let workbook_xml = read_zip_entry_string(&mut archive, "xl/workbook.xml");
    let sheet_rids = parse_workbook_sheet_rids(&workbook_xml);

    // Step 2: Read workbook rels to get rId → sheet file path
    let workbook_rels_xml = read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels");
    let rid_to_target = parse_rels_targets(&workbook_rels_xml);

    // Step 3: For each sheet, find its drawing and extract chart anchors
    let mut result: HashMap<String, Vec<(u32, Chart)>> = HashMap::new();

    for (sheet_name, sheet_rid) in &sheet_rids {
        let Some(sheet_target) = rid_to_target.get(sheet_rid) else {
            continue;
        };
        // Sheet target is relative to xl/ (e.g., "worksheets/sheet1.xml")
        let sheet_full_path = format!("xl/{sheet_target}");
        let sheet_filename = sheet_full_path.rsplit('/').next().unwrap_or(sheet_target);
        let sheet_rels_path = format!("xl/worksheets/_rels/{sheet_filename}.rels");

        let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path);
        if sheet_rels_xml.is_empty() {
            continue;
        }

        // Find drawing relationship
        let drawing_targets = parse_rels_by_type(&sheet_rels_xml, "drawing");
        for drawing_target in &drawing_targets {
            // Resolve relative path from worksheets/ to drawings/
            let drawing_path = resolve_relative_xl_path("xl/worksheets", drawing_target);
            let drawing_xml = read_zip_entry_string(&mut archive, &drawing_path);
            if drawing_xml.is_empty() {
                continue;
            }

            // Parse drawing for chart anchor positions
            let anchors = parse_drawing_chart_anchors(&drawing_xml);

            // Find drawing rels for chart rId resolution
            let drawing_filename = drawing_path.rsplit('/').next().unwrap_or(&drawing_path);
            let drawing_dir = drawing_path
                .rsplit_once('/')
                .map(|(d, _)| d)
                .unwrap_or("xl/drawings");
            let drawing_rels_path = format!("{drawing_dir}/_rels/{drawing_filename}.rels");
            let drawing_rels_xml = read_zip_entry_string(&mut archive, &drawing_rels_path);
            let drawing_rid_targets = parse_rels_targets(&drawing_rels_xml);

            for (anchor_row, chart_rid) in &anchors {
                let Some(chart_target) = drawing_rid_targets.get(chart_rid) else {
                    continue;
                };
                let chart_path = resolve_relative_xl_path(drawing_dir, chart_target);
                let chart_xml = read_zip_entry_string(&mut archive, &chart_path);
                if let Some(chart) = parse_chart_xml(&chart_xml) {
                    result
                        .entry(sheet_name.clone())
                        .or_default()
                        .push((*anchor_row, chart));
                }
            }
        }
    }

    // Step 4: Find any charts not associated with drawings (orphaned charts)
    // and assign them to the first sheet with u32::MAX sentinel
    let all_positioned_chart_paths: HashSet<String> = result
        .values()
        .flatten()
        .filter_map(|_| None::<String>) // We don't track paths, just check coverage below
        .collect();
    let _ = all_positioned_chart_paths; // consumed

    // Scan for chart XML files and check if they were captured by drawing anchors
    let chart_paths: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let entry = archive.by_index(i).ok()?;
            let name = entry.name().to_string();
            if name.starts_with("xl/charts/chart") && name.ends_with(".xml") {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    // Count total positioned charts
    let positioned_count: usize = result.values().map(|v| v.len()).sum();

    if chart_paths.len() > positioned_count {
        // Some charts weren't found via drawing anchors — parse them as unanchored
        let positioned_charts: HashSet<String> = collect_positioned_chart_paths(&result, data);

        let first_sheet = sheet_rids
            .first()
            .map(|(name, _)| name.clone())
            .unwrap_or_else(|| "Sheet1".to_string());

        for path in &chart_paths {
            if positioned_charts.contains(path) {
                continue;
            }
            let chart_xml = read_zip_entry_string(&mut archive, path);
            if let Some(chart) = parse_chart_xml(&chart_xml) {
                result
                    .entry(first_sheet.clone())
                    .or_default()
                    .push((u32::MAX, chart));
            }
        }
    }

    result
}

/// Collect the set of chart XML paths that were already positioned via drawing anchors.
pub(super) fn collect_positioned_chart_paths(
    chart_map: &HashMap<String, Vec<(u32, Chart)>>,
    data: &[u8],
) -> HashSet<String> {
    // Re-trace the drawing → chart resolution to find which chart paths are covered.
    // This is intentionally conservative — if we can't determine the path, we skip.
    let Ok(mut archive) = crate::parser::open_zip(data) else {
        return HashSet::new();
    };
    let mut positioned = HashSet::new();

    let workbook_xml = read_zip_entry_string(&mut archive, "xl/workbook.xml");
    let sheet_rids = parse_workbook_sheet_rids(&workbook_xml);
    let workbook_rels_xml = read_zip_entry_string(&mut archive, "xl/_rels/workbook.xml.rels");
    let rid_to_target = parse_rels_targets(&workbook_rels_xml);

    for (sheet_name, sheet_rid) in &sheet_rids {
        if !chart_map.contains_key(sheet_name) {
            continue;
        }
        let Some(sheet_target) = rid_to_target.get(sheet_rid) else {
            continue;
        };
        let sheet_full_path = format!("xl/{sheet_target}");
        let sheet_filename = sheet_full_path.rsplit('/').next().unwrap_or(sheet_target);
        let sheet_rels_path = format!("xl/worksheets/_rels/{sheet_filename}.rels");
        let sheet_rels_xml = read_zip_entry_string(&mut archive, &sheet_rels_path);
        let drawing_targets = parse_rels_by_type(&sheet_rels_xml, "drawing");

        for drawing_target in &drawing_targets {
            let drawing_path = resolve_relative_xl_path("xl/worksheets", drawing_target);
            let drawing_xml = read_zip_entry_string(&mut archive, &drawing_path);
            let anchors = parse_drawing_chart_anchors(&drawing_xml);
            let drawing_filename = drawing_path.rsplit('/').next().unwrap_or(&drawing_path);
            let drawing_dir = drawing_path
                .rsplit_once('/')
                .map(|(d, _)| d)
                .unwrap_or("xl/drawings");
            let drawing_rels_path = format!("{drawing_dir}/_rels/{drawing_filename}.rels");
            let drawing_rels_xml = read_zip_entry_string(&mut archive, &drawing_rels_path);
            let drawing_rid_targets = parse_rels_targets(&drawing_rels_xml);

            for (_row, chart_rid) in &anchors {
                if let Some(chart_target) = drawing_rid_targets.get(chart_rid) {
                    positioned.insert(resolve_relative_xl_path(drawing_dir, chart_target));
                }
            }
        }
    }

    positioned
}

/// Read a ZIP entry as a string. Returns empty string if not found.
pub(super) fn read_zip_entry_string(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    path: &str,
) -> String {
    let Ok(mut entry) = archive.by_name(path) else {
        return String::new();
    };
    let mut xml = String::new();
    let _ = std::io::Read::read_to_string(&mut entry, &mut xml);
    xml
}

/// Parse workbook.xml to extract sheet name → rId pairs (preserving order).
pub(super) fn parse_workbook_sheet_rids(xml: &str) -> Vec<(String, String)> {
    let mut result = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"sheet" {
                    let mut name = None;
                    let mut rid = None;
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"name" => {
                                if let Ok(v) = attr.unescape_value() {
                                    name = Some(v.to_string());
                                }
                            }
                            b"id" => {
                                if let Ok(v) = attr.unescape_value() {
                                    rid = Some(v.to_string());
                                }
                            }
                            _ => {}
                        }
                    }
                    if let (Some(n), Some(r)) = (name, rid) {
                        result.push((n, r));
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    result
}

/// Parse a .rels file to get Id → Target mapping.
pub(super) fn parse_rels_targets(xml: &str) -> HashMap<String, String> {
    xml_util::parse_rels_id_target(xml)
}

/// Parse a .rels file and return targets whose Type contains the given substring.
pub(super) fn parse_rels_by_type(xml: &str, type_substring: &str) -> Vec<String> {
    let mut targets = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let mut target = None;
                    let mut matches_type = false;
                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"Target" => {
                                if let Ok(v) = attr.unescape_value() {
                                    target = Some(v.to_string());
                                }
                            }
                            b"Type" => {
                                if let Ok(v) = attr.unescape_value()
                                    && v.contains(type_substring)
                                {
                                    matches_type = true;
                                }
                            }
                            _ => {}
                        }
                    }
                    if matches_type && let Some(t) = target {
                        targets.push(t);
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    targets
}

/// Resolve a relative path (like `../drawings/drawing1.xml`) against a base directory.
pub(super) fn resolve_relative_xl_path(base_dir: &str, relative: &str) -> String {
    xml_util::resolve_relative_path(base_dir, relative)
}

/// Parse drawing XML for chart anchor positions.
/// Returns (anchor_row, chart_rId) pairs from `<xdr:twoCellAnchor>` elements.
pub(super) fn parse_drawing_chart_anchors(xml: &str) -> Vec<(u32, String)> {
    let mut result = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    let mut in_two_cell_anchor = false;
    let mut in_from = false;
    let mut in_row = false;
    let mut anchor_row: Option<u32> = None;
    let mut chart_rid: Option<String> = None;
    let mut in_graphic_data = false;

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"twoCellAnchor" | b"oneCellAnchor" => {
                        in_two_cell_anchor = true;
                        anchor_row = None;
                        chart_rid = None;
                    }
                    b"from" if in_two_cell_anchor => {
                        in_from = true;
                    }
                    b"row" if in_from => {
                        in_row = true;
                    }
                    b"graphicData" if in_two_cell_anchor => {
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"uri"
                                && let Ok(val) = attr.unescape_value()
                                && val.contains("chart")
                            {
                                in_graphic_data = true;
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                if in_graphic_data && local.as_ref() == b"chart" {
                    for attr in e.attributes().flatten() {
                        if (attr.key.as_ref() == b"r:id" || attr.key.local_name().as_ref() == b"id")
                            && let Ok(val) = attr.unescape_value()
                        {
                            chart_rid = Some(val.to_string());
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Text(ref t)) => {
                if in_row
                    && let Ok(s) = t.xml_content()
                    && let Ok(row) = s.trim().parse::<u32>()
                {
                    anchor_row = Some(row);
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"twoCellAnchor" | b"oneCellAnchor" => {
                        if let (Some(row), Some(rid)) = (anchor_row.take(), chart_rid.take()) {
                            result.push((row, rid));
                        }
                        in_two_cell_anchor = false;
                        in_from = false;
                        in_graphic_data = false;
                    }
                    b"from" => {
                        in_from = false;
                    }
                    b"row" => {
                        in_row = false;
                    }
                    b"graphicData" => {
                        in_graphic_data = false;
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    result
}
