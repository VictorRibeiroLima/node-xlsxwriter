use std::collections::HashMap;

use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsBoolean, JsDate, JsNumber, JsObject, JsString, JsValue},
};
use rust_xlsxwriter::{
    Color, ConditionalFormat2ColorScale, ConditionalFormatType, ConditionalFormatValue,
};

use crate::node_xlsx::color::Color as NodeColor;

use crate::node_xlsx::util::js_date_to_naive_date_time;

use super::c_type::NodeXlsxConditionalFormatType;

struct NodeXlsxRule {
    pub r_type: ConditionalFormatType,
    pub value: ConditionalFormatValue,
}

impl NodeXlsxRule {
    pub fn from_js_value(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let r_type: Handle<JsString> = obj.get(cx, "type")?;
        let r_type = type_from_js_string(cx, r_type)?;
        let value: Handle<JsValue> = obj.get(cx, "value")?;
        let value = value_from_js_value(cx, value)?;
        Ok(Self { r_type, value })
    }
}

fn type_from_js_string(
    cx: &mut FunctionContext,
    value: Handle<JsString>,
) -> NeonResult<ConditionalFormatType> {
    let value = value.value(cx);
    match value.as_str() {
        "number" => Ok(ConditionalFormatType::Number),
        "percent" => Ok(ConditionalFormatType::Percent),
        "formula" => Ok(ConditionalFormatType::Formula),
        "percentile" => Ok(ConditionalFormatType::Percentile),
        "automatic" => Ok(ConditionalFormatType::Automatic),
        "lowest" => Ok(ConditionalFormatType::Lowest),
        "highest" => Ok(ConditionalFormatType::Highest),
        _ => {
            let err = format!("Invalid ConditionalFormatType: {}", value);
            cx.throw_error(err)
        }
    }
}

fn value_from_js_value(
    cx: &mut FunctionContext,
    value: Handle<JsValue>,
) -> NeonResult<ConditionalFormatValue> {
    if let Ok(string) = value.downcast::<JsString, _>(cx) {
        let string = string.value(cx);
        return Ok(string.into());
    } else if let Ok(number) = value.downcast::<JsNumber, _>(cx) {
        let number = number.value(cx);
        return Ok(number.into());
    } else if let Ok(date) = value.downcast::<JsDate, _>(cx) {
        let naive = &js_date_to_naive_date_time(cx, date)?;
        return Ok(naive.into());
    }
    let err = format!("Invalid ConditionalFormatValue: {:?}", value);
    cx.throw_error(err)
}
pub struct TwoColorScale {
    multi_range: Option<String>,
    stop_if_true: Option<bool>,
    min_color: Option<Color>,
    max_color: Option<Color>,
    min_rule: Option<NodeXlsxRule>,
    max_rule: Option<NodeXlsxRule>,
}

impl TwoColorScale {
    pub fn from_js_object(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let multi_range: Option<Handle<JsString>> = obj.get_opt(cx, "multiRange")?;
        let multi_range = multi_range.map(|range| range.value(cx));

        let stop_if_true: Option<Handle<JsBoolean>> = obj.get_opt(cx, "stopIfTrue")?;
        let stop_if_true = stop_if_true.map(|stop| stop.value(cx));

        let min_color: Option<Handle<JsObject>> = obj.get_opt(cx, "minColor")?;
        let min_color = match min_color {
            Some(color) => {
                let color = NodeColor::from_js_object(cx, color)?;
                let color = color.into();
                Some(color)
            }
            None => None,
        };

        let max_color: Option<Handle<JsObject>> = obj.get_opt(cx, "maxColor")?;
        let max_color = match max_color {
            Some(color) => {
                let color = NodeColor::from_js_object(cx, color)?;
                let color = color.into();
                Some(color)
            }
            None => None,
        };

        let min_rule: Option<Handle<JsObject>> = obj.get_opt(cx, "minRule")?;
        let min_rule = match min_rule {
            Some(rule) => Some(NodeXlsxRule::from_js_value(cx, rule)?),
            None => None,
        };

        let max_rule: Option<Handle<JsObject>> = obj.get_opt(cx, "maxRule")?;
        let max_rule = match max_rule {
            Some(rule) => Some(NodeXlsxRule::from_js_value(cx, rule)?),
            None => None,
        };

        Ok(Self {
            multi_range,
            stop_if_true,
            min_color,
            max_color,
            min_rule,
            max_rule,
        })
    }

    pub fn create_and_set_to_map(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &mut HashMap<u32, NodeXlsxConditionalFormatType>,
    ) -> NeonResult<u32> {
        let id: Handle<JsNumber> = obj.get(cx, "id")?;
        let id = id.value(cx) as u32;

        let two_color_scale = TwoColorScale::from_js_object(cx, obj)?;

        format_map.insert(
            id,
            NodeXlsxConditionalFormatType::TwoColorScale(two_color_scale.into()),
        );
        Ok(id)
    }
}

impl Into<ConditionalFormat2ColorScale> for TwoColorScale {
    fn into(self) -> ConditionalFormat2ColorScale {
        let mut scale = ConditionalFormat2ColorScale::new();
        if let Some(multi_range) = self.multi_range {
            scale = scale.set_multi_range(multi_range);
        }
        if let Some(stop_if_true) = self.stop_if_true {
            scale = scale.set_stop_if_true(stop_if_true);
        }
        if let Some(min_color) = self.min_color {
            scale = scale.set_minimum_color(min_color);
        }
        if let Some(max_color) = self.max_color {
            scale = scale.set_maximum_color(max_color);
        }
        if let Some(min_rule) = self.min_rule {
            scale = scale.set_minimum(min_rule.r_type, min_rule.value);
        }
        if let Some(max_rule) = self.max_rule {
            scale = scale.set_maximum(max_rule.r_type, max_rule.value);
        }
        scale
    }
}
