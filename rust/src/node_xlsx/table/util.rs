use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsObject, JsString},
};
use rust_xlsxwriter::{TableFunction, TableStyle};

use crate::node_xlsx::util::object_to_formula;

pub fn string_to_table_style(
    cx: &mut FunctionContext,
    js_str: Handle<JsString>,
) -> NeonResult<TableStyle> {
    let js_str = js_str.value(cx);
    match js_str.as_str() {
        "none" => Ok(TableStyle::None),
        "light1" => Ok(TableStyle::Light1),
        "light2" => Ok(TableStyle::Light2),
        "light3" => Ok(TableStyle::Light3),
        "light4" => Ok(TableStyle::Light4),
        "light5" => Ok(TableStyle::Light5),
        "light6" => Ok(TableStyle::Light6),
        "light7" => Ok(TableStyle::Light7),
        "light8" => Ok(TableStyle::Light8),
        "light9" => Ok(TableStyle::Light9),
        "light10" => Ok(TableStyle::Light10),
        "light11" => Ok(TableStyle::Light11),
        "light12" => Ok(TableStyle::Light12),
        "light13" => Ok(TableStyle::Light13),
        "light14" => Ok(TableStyle::Light14),
        "light15" => Ok(TableStyle::Light15),
        "light16" => Ok(TableStyle::Light16),
        "light17" => Ok(TableStyle::Light17),
        "light18" => Ok(TableStyle::Light18),
        "light19" => Ok(TableStyle::Light19),
        "light20" => Ok(TableStyle::Light20),
        "light21" => Ok(TableStyle::Light21),
        "medium1" => Ok(TableStyle::Medium1),
        "medium2" => Ok(TableStyle::Medium2),
        "medium3" => Ok(TableStyle::Medium3),
        "medium4" => Ok(TableStyle::Medium4),
        "medium5" => Ok(TableStyle::Medium5),
        "medium6" => Ok(TableStyle::Medium6),
        "medium7" => Ok(TableStyle::Medium7),
        "medium8" => Ok(TableStyle::Medium8),
        "medium9" => Ok(TableStyle::Medium9),
        "medium10" => Ok(TableStyle::Medium10),
        "medium11" => Ok(TableStyle::Medium11),
        "medium12" => Ok(TableStyle::Medium12),
        "medium13" => Ok(TableStyle::Medium13),
        "medium14" => Ok(TableStyle::Medium14),
        "medium15" => Ok(TableStyle::Medium15),
        "medium16" => Ok(TableStyle::Medium16),
        "medium17" => Ok(TableStyle::Medium17),
        "medium18" => Ok(TableStyle::Medium18),
        "medium19" => Ok(TableStyle::Medium19),
        "medium20" => Ok(TableStyle::Medium20),
        "medium21" => Ok(TableStyle::Medium21),
        "medium22" => Ok(TableStyle::Medium22),
        "medium23" => Ok(TableStyle::Medium23),
        "medium24" => Ok(TableStyle::Medium24),
        "medium25" => Ok(TableStyle::Medium25),
        "medium26" => Ok(TableStyle::Medium26),
        "medium27" => Ok(TableStyle::Medium27),
        "medium28" => Ok(TableStyle::Medium28),
        "dark1" => Ok(TableStyle::Dark1),
        "dark2" => Ok(TableStyle::Dark2),
        "dark3" => Ok(TableStyle::Dark3),
        "dark4" => Ok(TableStyle::Dark4),
        "dark5" => Ok(TableStyle::Dark5),
        "dark6" => Ok(TableStyle::Dark6),
        "dark7" => Ok(TableStyle::Dark7),
        "dark8" => Ok(TableStyle::Dark8),
        "dark9" => Ok(TableStyle::Dark9),
        "dark10" => Ok(TableStyle::Dark10),
        "dark11" => Ok(TableStyle::Dark11),
        _ => {
            let msg = format!("Unknown table style: {}", js_str);
            cx.throw_error(msg)
        }
    }
}

pub fn object_to_table_function(
    cx: &mut FunctionContext,
    js_obj: Handle<JsObject>,
) -> NeonResult<TableFunction> {
    let t_type: Handle<JsString> = js_obj.get(cx, "type")?;
    let formula: Option<Handle<JsObject>> = js_obj.get_opt(cx, "formula")?;

    let t_type = t_type.value(cx);

    match t_type.as_str() {
        "none" => Ok(TableFunction::None),
        "average" => Ok(TableFunction::Average),
        "count" => Ok(TableFunction::Count),
        "countNumbers" => Ok(TableFunction::CountNumbers),
        "max" => Ok(TableFunction::Max),
        "min" => Ok(TableFunction::Min),
        "sum" => Ok(TableFunction::Sum),
        "stdDev" => Ok(TableFunction::StdDev),
        "var" => Ok(TableFunction::Var),
        "custom" => {
            if let Some(formula) = formula {
                let (formula, _) = object_to_formula(cx, formula)?;
                Ok(TableFunction::Custom(formula))
            } else {
                cx.throw_error("Custom function requires a formula")
            }
        }
        _ => {
            let msg = format!("Unknown table function type: {}", t_type);
            cx.throw_error(msg)
        }
    }
}
