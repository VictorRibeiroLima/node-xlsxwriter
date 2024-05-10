/*
/**
 * @class ConditionalFormatDate
 * @classdesc Represents a Date style conditional format.
 * @extends ConditionalFormat
 * @property {Format} [format] - The format for the date.
 * @property {ConditionalFormatDateRule} rule - The rule for the date.
 * @property {string} [multiRange] - Is used to extend a conditional format over non-contiguous ranges like "B3:D6 I3:K6 B9:D12 I9:K12"
 * @property {boolean} [stopIfTrue] - Is used to set the “Stop if true” feature of a conditional formatting rule when more than one rule is applied to a cell or a range of cells. When this parameter is set then subsequent rules are not evaluated if the current rule is true.
 */ */

use std::collections::HashMap;

use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsBoolean, JsNumber, JsObject, JsString},
};
use rust_xlsxwriter::{ConditionalFormatDate, ConditionalFormatDateRule, Format};

use crate::node_xlsx::format::NodeXlsxFormat;

use super::c_type::NodeXlsxConditionalFormatType;

pub struct Date<'a> {
    format: Option<&'a Format>,
    rule: Option<ConditionalFormatDateRule>,
    multi_range: Option<String>,
    stop_if_true: Option<bool>,
}

impl<'a> Date<'a> {
    pub fn from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &'a mut HashMap<u32, Format>,
    ) -> NeonResult<Self> {
        let format: Option<Handle<JsObject>> = obj.get_opt(cx, "format")?;
        let format = match format {
            Some(format) => {
                let id: Handle<JsNumber> = format.get(cx, "id")?;
                let id = id.value(cx) as u32;
                if !format_map.contains_key(&id) {
                    let format = NodeXlsxFormat::from_js_object(cx, format)?;
                    format_map.insert(id, format.into());
                }
                let format = format_map.get(&id).unwrap();
                Some(format)
            }
            None => None,
        };

        let rule: Option<Handle<JsString>> = obj.get_opt(cx, "rule")?;
        let rule = rule.map(|rule| rule_from_js_string(cx, rule)).transpose()?;

        let multi_range: Option<Handle<JsString>> = obj.get_opt(cx, "multiRange")?;
        let multi_range = multi_range.map(|range| range.value(cx));

        let stop_if_true: Option<Handle<JsBoolean>> = obj.get_opt(cx, "stopIfTrue")?;
        let stop_if_true = stop_if_true.map(|stop| stop.value(cx));

        Ok(Self {
            format,
            rule,
            multi_range,
            stop_if_true,
        })
    }

    pub fn create_and_set_to_map(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        c_format_map: &mut HashMap<u32, NodeXlsxConditionalFormatType>,
        format_map: &'a mut HashMap<u32, Format>,
    ) -> NeonResult<u32> {
        let id: Handle<JsNumber> = obj.get(cx, "id")?;
        let id = id.value(cx) as u32;

        if !c_format_map.contains_key(&id) {
            let date = Date::from_js_object(cx, obj, format_map)?;
            let date = NodeXlsxConditionalFormatType::Date(date.into());
            c_format_map.insert(id, date);
        }

        Ok(id)
    }
}

impl Into<ConditionalFormatDate> for Date<'_> {
    fn into(self) -> ConditionalFormatDate {
        let mut date = ConditionalFormatDate::new();

        if let Some(format) = self.format {
            date = date.set_format(format);
        }

        if let Some(rule) = self.rule {
            date = date.set_rule(rule);
        }

        if let Some(multi_range) = self.multi_range {
            date = date.set_multi_range(multi_range);
        }

        if let Some(stop_if_true) = self.stop_if_true {
            date = date.set_stop_if_true(stop_if_true);
        }

        date
    }
}

fn rule_from_js_string(
    cx: &mut FunctionContext,
    string: Handle<JsString>,
) -> NeonResult<ConditionalFormatDateRule> {
    let string = string.value(cx);
    let rule = match string.as_str() {
        "yesterday" => ConditionalFormatDateRule::Yesterday,
        "today" => ConditionalFormatDateRule::Today,
        "tomorrow" => ConditionalFormatDateRule::Tomorrow,
        "last7Days" => ConditionalFormatDateRule::Last7Days,
        "lastWeek" => ConditionalFormatDateRule::LastWeek,
        "thisWeek" => ConditionalFormatDateRule::ThisWeek,
        "nextWeek" => ConditionalFormatDateRule::NextWeek,
        "lastMonth" => ConditionalFormatDateRule::LastMonth,
        "thisMonth" => ConditionalFormatDateRule::ThisMonth,
        "nextMonth" => ConditionalFormatDateRule::NextMonth,
        _ => {
            return cx.throw_error(format!("Invalid ConditionalFormatDateRule: {}", string));
        }
    };

    Ok(rule)
}
