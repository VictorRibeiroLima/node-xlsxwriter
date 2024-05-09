use std::collections::HashMap;

use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsBoolean, JsNumber, JsObject, JsString},
};
use rust_xlsxwriter::{
    ConditionalFormatDataBar, ConditionalFormatDataBarAxisPosition,
    ConditionalFormatDataBarDirection,
};

use crate::node_xlsx::color::Color;

use super::{c_type::NodeXlsxConditionalFormatType, rule::NodeXlsxFormatTypeRule};

pub struct DataBar {
    axis_color: Option<Color>,
    axis_position: Option<ConditionalFormatDataBarAxisPosition>,
    bar_only: Option<bool>,
    border_color: Option<Color>,
    border_off: Option<bool>,
    direction: Option<ConditionalFormatDataBarDirection>,
    fill_color: Option<Color>,
    max_rule: Option<NodeXlsxFormatTypeRule>,
    min_rule: Option<NodeXlsxFormatTypeRule>,
    negative_border_color: Option<Color>,
    negative_fill_color: Option<Color>,
    solid_fill: Option<bool>,
    multi_range: Option<String>,
    stop_if_true: Option<bool>,
}

impl DataBar {
    pub fn from_js_value(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let axis_color: Option<Handle<JsObject>> = obj.get_opt(cx, "axisColor")?;
        let axis_color = match axis_color {
            Some(axis_color) => Some(Color::from_js_object(cx, axis_color)?),
            None => None,
        };

        let axis_position: Option<Handle<JsString>> = obj.get_opt(cx, "axisPosition")?;
        let axis_position = match axis_position {
            Some(axis_position) => Some(axis_position_from_js_string(cx, axis_position)?),
            None => None,
        };

        let bar_only: Option<Handle<JsBoolean>> = obj.get_opt(cx, "barOnly")?;
        let bar_only = match bar_only {
            Some(bar_only) => Some(bar_only.value(cx)),
            None => None,
        };

        let border_color: Option<Handle<JsObject>> = obj.get_opt(cx, "borderColor")?;
        let border_color = match border_color {
            Some(border_color) => Some(Color::from_js_object(cx, border_color)?),
            None => None,
        };

        let border_off: Option<Handle<JsBoolean>> = obj.get_opt(cx, "borderOff")?;
        let border_off = match border_off {
            Some(border_off) => Some(border_off.value(cx)),
            None => None,
        };

        let direction: Option<Handle<JsString>> = obj.get_opt(cx, "direction")?;
        let direction = match direction {
            Some(direction) => Some(bar_direction_from_js_string(cx, direction)?),
            None => None,
        };

        let fill_color: Option<Handle<JsObject>> = obj.get_opt(cx, "fillColor")?;
        let fill_color = match fill_color {
            Some(fill_color) => Some(Color::from_js_object(cx, fill_color)?),
            None => None,
        };

        let max_rule: Option<Handle<JsObject>> = obj.get_opt(cx, "maxRule")?;
        let max_rule = match max_rule {
            Some(max_rule) => Some(NodeXlsxFormatTypeRule::from_js_value(cx, max_rule)?),
            None => None,
        };

        let min_rule: Option<Handle<JsObject>> = obj.get_opt(cx, "minRule")?;
        let min_rule = match min_rule {
            Some(min_rule) => Some(NodeXlsxFormatTypeRule::from_js_value(cx, min_rule)?),
            None => None,
        };

        let negative_border_color: Option<Handle<JsObject>> =
            obj.get_opt(cx, "negativeBorderColor")?;
        let negative_border_color = match negative_border_color {
            Some(negative_border_color) => Some(Color::from_js_object(cx, negative_border_color)?),
            None => None,
        };

        let negative_fill_color: Option<Handle<JsObject>> = obj.get_opt(cx, "negativeFillColor")?;
        let negative_fill_color = match negative_fill_color {
            Some(negative_fill_color) => Some(Color::from_js_object(cx, negative_fill_color)?),
            None => None,
        };

        let solid_fill: Option<Handle<JsBoolean>> = obj.get_opt(cx, "solidFill")?;
        let solid_fill = match solid_fill {
            Some(solid_fill) => Some(solid_fill.value(cx)),
            None => None,
        };

        let multi_range: Option<Handle<JsString>> = obj.get_opt(cx, "multiRange")?;
        let multi_range = match multi_range {
            Some(multi_range) => Some(multi_range.value(cx)),
            None => None,
        };

        let stop_if_true: Option<Handle<JsBoolean>> = obj.get_opt(cx, "stopIfTrue")?;
        let stop_if_true = match stop_if_true {
            Some(stop_if_true) => Some(stop_if_true.value(cx)),
            None => None,
        };

        Ok(Self {
            axis_color,
            axis_position,
            bar_only,
            border_color,
            border_off,
            direction,
            fill_color,
            max_rule,
            min_rule,
            negative_border_color,
            negative_fill_color,
            solid_fill,
            multi_range,
            stop_if_true,
        })
    }

    pub fn create_and_set_to_map(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        c_format_map: &mut HashMap<u32, NodeXlsxConditionalFormatType>,
    ) -> NeonResult<u32> {
        let id: Handle<JsNumber> = obj.get(cx, "id")?;
        let id = id.value(cx) as u32;

        if !c_format_map.contains_key(&id) {
            let format_data_bar = DataBar::from_js_value(cx, obj)?;
            let data_bar: ConditionalFormatDataBar = format_data_bar.into();

            c_format_map.insert(id, NodeXlsxConditionalFormatType::DataBar(data_bar));
        }

        Ok(id)
    }
}

impl Into<ConditionalFormatDataBar> for DataBar {
    fn into(self) -> ConditionalFormatDataBar {
        let mut data_bar = ConditionalFormatDataBar::new();

        if let Some(axis_color) = self.axis_color {
            let color: rust_xlsxwriter::Color = axis_color.into();
            data_bar = data_bar.set_axis_color(color);
        };

        if let Some(axis_position) = self.axis_position {
            data_bar = data_bar.set_axis_position(axis_position);
        };

        if let Some(bar_only) = self.bar_only {
            data_bar = data_bar.set_bar_only(bar_only);
        };

        if let Some(border_color) = self.border_color {
            let color: rust_xlsxwriter::Color = border_color.into();
            data_bar = data_bar.set_border_color(color);
        };

        if let Some(border_off) = self.border_off {
            data_bar = data_bar.set_border_off(border_off);
        };

        if let Some(direction) = self.direction {
            data_bar = data_bar.set_direction(direction);
        };

        if let Some(fill_color) = self.fill_color {
            let color: rust_xlsxwriter::Color = fill_color.into();
            data_bar = data_bar.set_fill_color(color);
        };

        if let Some(max_rule) = self.max_rule {
            data_bar = data_bar.set_maximum(max_rule.r_type, max_rule.value);
        };

        if let Some(min_rule) = self.min_rule {
            data_bar = data_bar.set_minimum(min_rule.r_type, min_rule.value);
        };

        if let Some(negative_border_color) = self.negative_border_color {
            let color: rust_xlsxwriter::Color = negative_border_color.into();
            data_bar = data_bar.set_negative_border_color(color);
        };

        if let Some(negative_fill_color) = self.negative_fill_color {
            let color: rust_xlsxwriter::Color = negative_fill_color.into();
            data_bar = data_bar.set_negative_fill_color(color);
        };

        if let Some(solid_fill) = self.solid_fill {
            data_bar = data_bar.set_solid_fill(solid_fill);
        };

        if let Some(multi_range) = self.multi_range {
            data_bar = data_bar.set_multi_range(multi_range);
        };

        if let Some(stop_if_true) = self.stop_if_true {
            data_bar = data_bar.set_stop_if_true(stop_if_true);
        };

        data_bar
    }
}

fn axis_position_from_js_string(
    cx: &mut FunctionContext,
    value: Handle<JsString>,
) -> NeonResult<ConditionalFormatDataBarAxisPosition> {
    let value = value.value(cx);
    match value.as_str() {
        "automatic" => Ok(ConditionalFormatDataBarAxisPosition::Automatic),
        "midPoint" => Ok(ConditionalFormatDataBarAxisPosition::Midpoint),
        "none" => Ok(ConditionalFormatDataBarAxisPosition::None),
        _ => cx.throw_error(format!("Invalid value for axis position: {}", value)),
    }
}

fn bar_direction_from_js_string(
    cx: &mut FunctionContext,
    value: Handle<JsString>,
) -> NeonResult<ConditionalFormatDataBarDirection> {
    let value = value.value(cx);
    match value.as_str() {
        "context" => Ok(ConditionalFormatDataBarDirection::Context),
        "leftToRight" => Ok(ConditionalFormatDataBarDirection::LeftToRight),
        "rightToLeft" => Ok(ConditionalFormatDataBarDirection::RightToLeft),
        _ => cx.throw_error(format!("Invalid value for bar direction: {}", value)),
    }
}
