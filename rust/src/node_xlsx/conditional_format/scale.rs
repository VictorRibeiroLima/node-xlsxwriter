use std::collections::HashMap;

use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsBoolean, JsNumber, JsObject, JsString},
};
use rust_xlsxwriter::{Color, ConditionalFormat2ColorScale, ConditionalFormat3ColorScale};

use crate::node_xlsx::color::Color as NodeColor;

use super::{c_type::NodeXlsxConditionalFormatType, rule::NodeXlsxRule};

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

pub struct ThreeColorScale {
    mid_color: Option<Color>,
    mid_rule: Option<NodeXlsxRule>,
    two_color_scale: TwoColorScale,
}

impl ThreeColorScale {
    pub fn from_js_object(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let mid_color: Option<Handle<JsObject>> = obj.get_opt(cx, "midColor")?;
        let mid_color = match mid_color {
            Some(color) => {
                let color = NodeColor::from_js_object(cx, color)?;
                let color = color.into();
                Some(color)
            }
            None => None,
        };

        let mid_rule: Option<Handle<JsObject>> = obj.get_opt(cx, "midRule")?;
        let mid_rule = match mid_rule {
            Some(rule) => Some(NodeXlsxRule::from_js_value(cx, rule)?),
            None => None,
        };

        let two_color_scale = TwoColorScale::from_js_object(cx, obj)?;

        Ok(Self {
            mid_color,
            mid_rule,
            two_color_scale,
        })
    }

    pub fn create_and_set_to_map(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &mut HashMap<u32, NodeXlsxConditionalFormatType>,
    ) -> NeonResult<u32> {
        let id: Handle<JsNumber> = obj.get(cx, "id")?;
        let id = id.value(cx) as u32;

        let three_color_scale = ThreeColorScale::from_js_object(cx, obj)?;

        format_map.insert(
            id,
            NodeXlsxConditionalFormatType::ThreeColorScale(three_color_scale.into()),
        );
        Ok(id)
    }
}

impl Into<ConditionalFormat3ColorScale> for ThreeColorScale {
    fn into(self) -> ConditionalFormat3ColorScale {
        let mut scale = ConditionalFormat3ColorScale::new();
        if let Some(mid_color) = self.mid_color {
            scale = scale.set_midpoint_color(mid_color);
        }
        if let Some(mid_rule) = self.mid_rule {
            scale = scale.set_midpoint(mid_rule.r_type, mid_rule.value);
        }
        if let Some(multi_range) = self.two_color_scale.multi_range {
            scale = scale.set_multi_range(multi_range);
        }
        if let Some(stop_if_true) = self.two_color_scale.stop_if_true {
            scale = scale.set_stop_if_true(stop_if_true);
        }
        if let Some(min_color) = self.two_color_scale.min_color {
            scale = scale.set_minimum_color(min_color);
        }
        if let Some(max_color) = self.two_color_scale.max_color {
            scale = scale.set_maximum_color(max_color);
        }
        if let Some(min_rule) = self.two_color_scale.min_rule {
            scale = scale.set_minimum(min_rule.r_type, min_rule.value);
        }
        if let Some(max_rule) = self.two_color_scale.max_rule {
            scale = scale.set_maximum(max_rule.r_type, max_rule.value);
        }
        scale
    }
}
