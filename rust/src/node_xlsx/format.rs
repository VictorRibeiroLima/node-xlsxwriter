use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsBoolean, JsNumber, JsObject, JsString},
};
use rust_xlsxwriter::{FormatAlign, FormatPattern, FormatUnderline};

use super::{
    border::{Border, DiagonalBorder},
    color::Color,
};

pub struct Format {
    align: Option<FormatAlign>,
    background_color: Option<Color>,
    bold: Option<bool>,
    left_border: Option<Border>,
    right_border: Option<Border>,
    top_border: Option<Border>,
    bottom_border: Option<Border>,
    diagonal_border: Option<DiagonalBorder>,
    charset: Option<u8>,
    font_color: Option<Color>,
    font_family: Option<u8>,
    font_name: Option<String>,
    font_scheme: Option<String>,
    font_size: Option<u32>,
    strike_through: Option<bool>,
    foreground_color: Option<Color>,
    hidden: Option<bool>,
    hyperlink: Option<bool>,
    indent: Option<u8>,
    italic: Option<bool>,
    locked: Option<bool>,
    num_fmt: Option<String>,
    num_fmt_id: Option<u8>,
    pattern: Option<FormatPattern>,
    underline: Option<FormatUnderline>,
}

impl Format {
    pub fn from_js_object(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let result = Self::inner_from_object(cx, &obj);
        match result {
            Ok(format) => Ok(format),
            Err(error) => {
                let error = format!("Error parsing Format: {:?} with error:\n  {}", obj, error);
                let js_string = cx.string(error);
                cx.throw(js_string)
            }
        }
    }

    fn inner_from_object(cx: &mut FunctionContext, object: &Handle<JsObject>) -> NeonResult<Self> {
        let align: Option<Handle<JsString>> = object.get_opt(cx, "align")?;
        let align: Option<FormatAlign> = match align {
            Some(align) => Some(align_from_js_string(cx, align)?),
            None => None,
        };

        let background_color: Option<Handle<JsObject>> = object.get_opt(cx, "backgroundColor")?;
        let background_color: Option<Color> = match background_color {
            Some(background_color) => Some(Color::from_js_object(cx, background_color)?),
            None => None,
        };

        let bold: Option<Handle<JsBoolean>> = object.get_opt(cx, "bold")?;
        let bold = match bold {
            Some(bold) => Some(bold.value(cx)),
            None => None,
        };

        let left_border: Option<Handle<JsObject>> = object.get_opt(cx, "leftBorder")?;
        let left_border = match left_border {
            Some(left_border) => Some(Border::from_js_object(cx, left_border)?),
            None => None,
        };

        let right_border: Option<Handle<JsObject>> = object.get_opt(cx, "rightBorder")?;
        let right_border = match right_border {
            Some(right_border) => Some(Border::from_js_object(cx, right_border)?),
            None => None,
        };

        let top_border: Option<Handle<JsObject>> = object.get_opt(cx, "topBorder")?;
        let top_border = match top_border {
            Some(top_border) => Some(Border::from_js_object(cx, top_border)?),
            None => None,
        };

        let bottom_border: Option<Handle<JsObject>> = object.get_opt(cx, "bottomBorder")?;
        let bottom_border = match bottom_border {
            Some(bottom_border) => Some(Border::from_js_object(cx, bottom_border)?),
            None => None,
        };

        let diagonal_border: Option<Handle<JsObject>> = object.get_opt(cx, "diagonalBorder")?;
        let diagonal_border = match diagonal_border {
            Some(diagonal_border) => Some(DiagonalBorder::from_js_object(cx, diagonal_border)?),
            None => None,
        };

        let charset: Option<Handle<JsNumber>> = object.get_opt(cx, "charset")?;
        let charset = match charset {
            Some(charset) => Some(charset.value(cx) as u8),
            None => None,
        };

        let font_color: Option<Handle<JsObject>> = object.get_opt(cx, "fontColor")?;
        let font_color = match font_color {
            Some(font_color) => Some(Color::from_js_object(cx, font_color)?),
            None => None,
        };

        let font_family: Option<Handle<JsNumber>> = object.get_opt(cx, "fontFamily")?;
        let font_family = match font_family {
            Some(font_family) => Some(font_family.value(cx) as u8),
            None => None,
        };

        let font_name: Option<Handle<JsString>> = object.get_opt(cx, "fontName")?;
        let font_name = match font_name {
            Some(font_name) => Some(font_name.value(cx)),
            None => None,
        };

        let font_scheme: Option<Handle<JsString>> = object.get_opt(cx, "fontScheme")?;
        let font_scheme = match font_scheme {
            Some(font_scheme) => Some(font_scheme.value(cx)),
            None => None,
        };

        let font_size: Option<Handle<JsNumber>> = object.get_opt(cx, "fontSize")?;
        let font_size = match font_size {
            Some(font_size) => Some(font_size.value(cx) as u32),
            None => None,
        };

        let strike_through: Option<Handle<JsBoolean>> = object.get_opt(cx, "strikeThrough")?;
        let strike_through = match strike_through {
            Some(strike_through) => Some(strike_through.value(cx)),
            None => None,
        };

        let foreground_color: Option<Handle<JsObject>> = object.get_opt(cx, "foregroundColor")?;
        let foreground_color = match foreground_color {
            Some(foreground_color) => Some(Color::from_js_object(cx, foreground_color)?),
            None => None,
        };

        let hidden: Option<Handle<JsBoolean>> = object.get_opt(cx, "hidden")?;
        let hidden = match hidden {
            Some(hidden) => Some(hidden.value(cx)),
            None => None,
        };

        let hyperlink: Option<Handle<JsBoolean>> = object.get_opt(cx, "hyperlink")?;
        let hyperlink = match hyperlink {
            Some(hyperlink) => Some(hyperlink.value(cx)),
            None => None,
        };

        let indent: Option<Handle<JsNumber>> = object.get_opt(cx, "indent")?;
        let indent = match indent {
            Some(indent) => Some(indent.value(cx) as u8),
            None => None,
        };

        let italic: Option<Handle<JsBoolean>> = object.get_opt(cx, "italic")?;
        let italic = match italic {
            Some(italic) => Some(italic.value(cx)),
            None => None,
        };

        let locked: Option<Handle<JsBoolean>> = object.get_opt(cx, "locked")?;
        let locked = match locked {
            Some(locked) => Some(locked.value(cx)),
            None => None,
        };

        let num_fmt: Option<Handle<JsString>> = object.get_opt(cx, "numFmt")?;
        let num_fmt = match num_fmt {
            Some(num_fmt) => Some(num_fmt.value(cx)),
            None => None,
        };

        let num_fmt_id: Option<Handle<JsNumber>> = object.get_opt(cx, "numFmtId")?;
        let num_fmt_id = match num_fmt_id {
            Some(num_fmt_id) => Some(num_fmt_id.value(cx) as u8),
            None => None,
        };

        let pattern: Option<Handle<JsString>> = object.get_opt(cx, "pattern")?;
        let pattern = match pattern {
            Some(pattern) => Some(pattern_from_js_string(cx, pattern)?),
            None => None,
        };

        let underline: Option<Handle<JsString>> = object.get_opt(cx, "underline")?;
        let underline = match underline {
            Some(underline) => Some(underline_from_js_string(cx, underline)?),
            None => None,
        };

        Ok(Self {
            align,
            background_color,
            bold,
            left_border,
            right_border,
            top_border,
            bottom_border,
            diagonal_border,
            charset,
            font_color,
            font_family,
            font_name,
            font_scheme,
            font_size,
            strike_through,
            foreground_color,
            hidden,
            hyperlink,
            indent,
            italic,
            locked,
            num_fmt,
            num_fmt_id,
            pattern,
            underline,
        })
    }
}

impl Into<rust_xlsxwriter::Format> for Format {
    fn into(self) -> rust_xlsxwriter::Format {
        let mut format = rust_xlsxwriter::Format::new();

        if let Some(align) = self.align {
            format = format.set_align(align);
        }

        if let Some(background_color) = self.background_color {
            let color: rust_xlsxwriter::Color = background_color.into();
            format = format.set_background_color(color);
        }

        if let Some(bold) = self.bold {
            if bold {
                format = format.set_bold();
            }
        }

        if let Some(left_border) = self.left_border {
            let color: rust_xlsxwriter::Color = left_border.color.into();
            format = format
                .set_border_left(left_border.b_type)
                .set_border_left_color(color);
        }

        if let Some(right_border) = self.right_border {
            let color: rust_xlsxwriter::Color = right_border.color.into();
            format = format
                .set_border_right(right_border.b_type)
                .set_border_right_color(color);
        }

        if let Some(top_border) = self.top_border {
            let color: rust_xlsxwriter::Color = top_border.color.into();
            format = format
                .set_border_top(top_border.b_type)
                .set_border_top_color(color);
        }

        if let Some(bottom_border) = self.bottom_border {
            let color: rust_xlsxwriter::Color = bottom_border.color.into();
            format = format
                .set_border_bottom(bottom_border.b_type)
                .set_border_bottom_color(color);
        }

        if let Some(diagonal_border) = self.diagonal_border {
            let color: rust_xlsxwriter::Color = diagonal_border.border.color.into();
            format = format
                .set_border_diagonal(diagonal_border.border.b_type)
                .set_border_diagonal_color(color)
                .set_border_diagonal_type(diagonal_border.d_type);
        }

        if let Some(charset) = self.charset {
            format = format.set_font_charset(charset);
        }

        if let Some(font_color) = self.font_color {
            let color: rust_xlsxwriter::Color = font_color.into();
            format = format.set_font_color(color);
        }

        if let Some(font_family) = self.font_family {
            format = format.set_font_family(font_family);
        }

        if let Some(font_name) = self.font_name {
            format = format.set_font_name(&font_name);
        }

        if let Some(font_scheme) = self.font_scheme {
            format = format.set_font_scheme(&font_scheme);
        }

        if let Some(font_size) = self.font_size {
            format = format.set_font_size(font_size);
        }

        if let Some(strike_through) = self.strike_through {
            if strike_through {
                format = format.set_font_strikethrough();
            }
        }

        if let Some(foreground_color) = self.foreground_color {
            let color: rust_xlsxwriter::Color = foreground_color.into();
            format = format.set_foreground_color(color);
        }

        if let Some(hidden) = self.hidden {
            if hidden {
                format = format.set_hidden();
            }
        }

        if let Some(hyperlink) = self.hyperlink {
            if hyperlink {
                format = format.set_hyperlink();
            }
        }

        if let Some(indent) = self.indent {
            format = format.set_indent(indent);
        }

        if let Some(italic) = self.italic {
            if italic {
                format = format.set_italic();
            }
        }

        if let Some(locked) = self.locked {
            if locked {
                format = format.set_locked();
            }
        }

        if let Some(num_fmt) = self.num_fmt {
            format = format.set_num_format(&num_fmt);
        }

        if let Some(num_fmt_id) = self.num_fmt_id {
            format = format.set_num_format_index(num_fmt_id);
        }

        if let Some(pattern) = self.pattern {
            format = format.set_pattern(pattern);
        }

        if let Some(underline) = self.underline {
            format = format.set_underline(underline);
        }

        return format;
    }
}

fn align_from_js_string(
    cx: &mut FunctionContext,
    js_string: Handle<JsString>,
) -> NeonResult<FormatAlign> {
    let js_string = js_string.value(cx);
    match js_string.as_str() {
        "general" => Ok(FormatAlign::General),
        "left" => Ok(FormatAlign::Left),
        "center" => Ok(FormatAlign::Center),
        "right" => Ok(FormatAlign::Right),
        "fill" => Ok(FormatAlign::Fill),
        "justify" => Ok(FormatAlign::Justify),
        "centerAcross" => Ok(FormatAlign::CenterAcross),
        "distributed" => Ok(FormatAlign::Distributed),
        "top" => Ok(FormatAlign::Top),
        "bottom" => Ok(FormatAlign::Bottom),
        "verticalCenter" => Ok(FormatAlign::VerticalCenter),
        "verticalJustify" => Ok(FormatAlign::VerticalJustify),
        "verticalDistributed" => Ok(FormatAlign::VerticalDistributed),
        _ => {
            let error = format!("Unknown align type: {}", js_string);
            let js_string = cx.string(&error);
            cx.throw(js_string)
        }
    }
}

fn pattern_from_js_string(
    cx: &mut FunctionContext,
    js_string: Handle<JsString>,
) -> NeonResult<FormatPattern> {
    let js_string = js_string.value(cx);
    match js_string.as_str() {
        "none" => Ok(FormatPattern::None),
        "solid" => Ok(FormatPattern::Solid),
        "mediumGray" => Ok(FormatPattern::MediumGray),
        "darkGray" => Ok(FormatPattern::DarkGray),
        "lightGray" => Ok(FormatPattern::LightGray),
        "darkHorizontal" => Ok(FormatPattern::DarkHorizontal),
        "darkVertical" => Ok(FormatPattern::DarkVertical),
        "darkDown" => Ok(FormatPattern::DarkDown),
        "darkUp" => Ok(FormatPattern::DarkUp),
        "darkGrid" => Ok(FormatPattern::DarkGrid),
        "darkTrellis" => Ok(FormatPattern::DarkTrellis),
        "lightHorizontal" => Ok(FormatPattern::LightHorizontal),
        "lightVertical" => Ok(FormatPattern::LightVertical),
        "lightDown" => Ok(FormatPattern::LightDown),
        "lightUp" => Ok(FormatPattern::LightUp),
        "lightGrid" => Ok(FormatPattern::LightGrid),
        "lightTrellis" => Ok(FormatPattern::LightTrellis),
        "gray125" => Ok(FormatPattern::Gray125),
        "gray0625" => Ok(FormatPattern::Gray0625),
        _ => {
            let error = format!("Unknown pattern type: {}", js_string);
            let js_string = cx.string(&error);
            cx.throw(js_string)
        }
    }
}

fn underline_from_js_string(
    cx: &mut FunctionContext,
    js_string: Handle<JsString>,
) -> NeonResult<FormatUnderline> {
    let js_string = js_string.value(cx);
    match js_string.as_str() {
        "none" => Ok(FormatUnderline::None),
        "single" => Ok(FormatUnderline::Single),
        "double" => Ok(FormatUnderline::Double),
        "singleAccounting" => Ok(FormatUnderline::SingleAccounting),
        "doubleAccounting" => Ok(FormatUnderline::DoubleAccounting),
        _ => {
            let error = format!("Unknown underline type: {}", js_string);
            let js_string = cx.string(&error);
            cx.throw(js_string)
        }
    }
}
