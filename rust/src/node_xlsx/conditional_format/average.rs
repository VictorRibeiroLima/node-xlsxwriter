use std::collections::HashMap;

use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsBoolean, JsNumber, JsObject, JsString},
};
use rust_xlsxwriter::{ConditionalFormatAverage, ConditionalFormatAverageRule, Format};

use crate::node_xlsx::format::NodeXlsxFormat;

use super::c_type::NodeXlsxConditionalFormatType;

pub struct Average<'a> {
    multi_range: Option<String>,
    stop_if_true: Option<bool>,
    rule: Option<ConditionalFormatAverageRule>,
    format: Option<&'a Format>,
}

impl<'a> Average<'a> {
    pub fn from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &'a mut HashMap<u32, Format>,
    ) -> NeonResult<Self> {
        let multi_range: Option<Handle<JsString>> = obj.get_opt(cx, "multiRange")?;
        let multi_range = multi_range.map(|range| range.value(cx));

        let stop_if_true: Option<Handle<JsBoolean>> = obj.get_opt(cx, "stopIfTrue")?;
        let stop_if_true = stop_if_true.map(|stop| stop.value(cx));

        let rule: Option<Handle<JsString>> = obj.get_opt(cx, "rule")?;
        let rule = match rule {
            Some(rule) => Some(average_rule_from_js_string(cx, rule)?),
            None => None,
        };

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

        Ok(Self {
            multi_range,
            stop_if_true,
            rule,
            format,
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

        let average = Average::from_js_object(cx, obj, format_map)?;

        c_format_map.insert(id, NodeXlsxConditionalFormatType::Average(average.into()));
        Ok(id)
    }
}

impl Into<ConditionalFormatAverage> for Average<'_> {
    fn into(self) -> ConditionalFormatAverage {
        let mut average = ConditionalFormatAverage::new();
        if let Some(multi_range) = self.multi_range {
            average = average.set_multi_range(multi_range);
        }
        if let Some(stop_if_true) = self.stop_if_true {
            average = average.set_stop_if_true(stop_if_true);
        }
        if let Some(rule) = self.rule {
            average = average.set_rule(rule);
        }
        if let Some(format) = self.format {
            average = average.set_format(format);
        }
        average
    }
}

fn average_rule_from_js_string(
    cx: &mut FunctionContext,
    value: Handle<JsString>,
) -> NeonResult<ConditionalFormatAverageRule> {
    let value = value.value(cx);
    match value.as_str() {
        "aboveAverage" => Ok(ConditionalFormatAverageRule::AboveAverage),
        "belowAverage" => Ok(ConditionalFormatAverageRule::BelowAverage),
        "equalOrAboveAverage" => Ok(ConditionalFormatAverageRule::EqualOrAboveAverage),
        "equalOrBelowAverage" => Ok(ConditionalFormatAverageRule::EqualOrBelowAverage),
        "oneStandardDeviationAbove" => Ok(ConditionalFormatAverageRule::OneStandardDeviationAbove),
        "oneStandardDeviationBelow" => Ok(ConditionalFormatAverageRule::OneStandardDeviationBelow),
        "twoStandardDeviationsAbove" => {
            Ok(ConditionalFormatAverageRule::TwoStandardDeviationsAbove)
        }
        "twoStandardDeviationsBelow" => {
            Ok(ConditionalFormatAverageRule::TwoStandardDeviationsBelow)
        }
        "threeStandardDeviationsAbove" => {
            Ok(ConditionalFormatAverageRule::ThreeStandardDeviationsAbove)
        }
        "threeStandardDeviationsBelow" => {
            Ok(ConditionalFormatAverageRule::ThreeStandardDeviationsBelow)
        }

        _ => {
            let err = format!("Invalid ConditionalFormatAverageRule: {}", value);
            cx.throw_error(err)
        }
    }
}
