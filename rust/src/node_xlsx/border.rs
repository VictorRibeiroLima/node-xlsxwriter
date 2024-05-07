use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsObject, JsString},
};
use rust_xlsxwriter::{FormatBorder, FormatDiagonalBorder};

use super::color::Color;

#[derive(Clone, Copy)]
pub struct Border {
    pub b_type: FormatBorder,
    pub color: Color,
}

impl Border {
    pub fn from_js_object(cx: &mut FunctionContext, object: Handle<JsObject>) -> NeonResult<Self> {
        let b_type = object.get(cx, "style")?;
        let color = object.get(cx, "color")?;
        Ok(Self {
            b_type: format_border_from_js_string(cx, b_type)?,
            color: Color::from_js_object(cx, color)?,
        })
    }
}

pub struct DiagonalBorder {
    pub border: Border,
    pub d_type: FormatDiagonalBorder,
}

impl DiagonalBorder {
    pub fn from_js_object(cx: &mut FunctionContext, object: Handle<JsObject>) -> NeonResult<Self> {
        let b_type = object.get(cx, "style")?;
        let color = object.get(cx, "color")?;
        let d_style = object.get(cx, "dStyle")?;
        let border = Border {
            b_type: format_border_from_js_string(cx, b_type)?,
            color: Color::from_js_object(cx, color)?,
        };
        Ok(Self {
            border,
            d_type: diagonal_border_from_js_string(cx, d_style)?,
        })
    }
}

fn format_border_from_js_string(
    cx: &mut FunctionContext,
    border: Handle<JsString>,
) -> NeonResult<FormatBorder> {
    let border_type = border.value(cx);
    match border_type.as_str() {
        "none" => Ok(FormatBorder::None),
        "thin" => Ok(FormatBorder::Thin),
        "medium" => Ok(FormatBorder::Medium),
        "thick" => Ok(FormatBorder::Thick),
        "double" => Ok(FormatBorder::Double),
        "hair" => Ok(FormatBorder::Hair),
        "medium_dashed" => Ok(FormatBorder::MediumDashed),
        "dash_dot" => Ok(FormatBorder::DashDot),
        "medium_dash_dot" => Ok(FormatBorder::MediumDashDot),
        "dash_dot_dot" => Ok(FormatBorder::DashDotDot),
        "medium_dash_dot_dot" => Ok(FormatBorder::MediumDashDotDot),
        "slant_dash_dot" => Ok(FormatBorder::SlantDashDot),
        _ => {
            let error = format!("Unknown border type: {}", border_type);
            let js_string = cx.string(error);
            cx.throw(js_string)
        }
    }
}

fn diagonal_border_from_js_string(
    cx: &mut FunctionContext,
    border: Handle<JsString>,
) -> NeonResult<FormatDiagonalBorder> {
    let border_type = border.value(cx);
    match border_type.as_str() {
        "none" => Ok(FormatDiagonalBorder::None),
        "borderUp" => Ok(FormatDiagonalBorder::BorderUp),
        "borderDown" => Ok(FormatDiagonalBorder::BorderDown),
        "borderUpDown" => Ok(FormatDiagonalBorder::BorderUpDown),
        _ => {
            let error = format!("Unknown diagonal border type: {}", border_type);
            let js_string = cx.string(error);
            cx.throw(js_string)
        }
    }
}
