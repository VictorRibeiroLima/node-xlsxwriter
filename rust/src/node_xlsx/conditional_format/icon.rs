use std::collections::HashMap;

use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsArray, JsBoolean, JsNumber, JsObject, JsString},
};
use rust_xlsxwriter::{
    ConditionalFormatCustomIcon, ConditionalFormatIconSet, ConditionalFormatIconType,
};

use super::{c_type::NodeXlsxConditionalFormatType, rule::NodeXlsxFormatTypeRule};

struct CustomIcon {
    pub greater_than: bool,
    pub no_icon: bool,
    pub i_type: Option<ConditionalFormatIconType>,
    pub i_type_index: Option<u8>,
    pub rule: Option<NodeXlsxFormatTypeRule>,
}

impl CustomIcon {
    pub fn from_js_object(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let greater_than: Handle<JsBoolean> = obj.get(cx, "greaterThan")?;
        let greater_than = greater_than.value(cx);
        let mut i_type_index: Option<u8> = None;

        let no_icon: Handle<JsBoolean> = obj.get(cx, "noIcon")?;
        let no_icon = no_icon.value(cx);

        let i_type: Option<Handle<JsString>> = obj.get_opt(cx, "iconType")?;
        let i_type = match i_type {
            Some(i_type) => {
                let index: Option<Handle<JsNumber>> = obj.get_opt(cx, "iconTypeIndex")?;
                let index = match index {
                    Some(index) => Some(index.value(cx) as u8),
                    None => {
                        let error = format!("iconTypeIndex is required when icon is set");
                        return cx.throw_error(error);
                    }
                };
                i_type_index = index;
                Some(icon_from_js_string(cx, i_type)?)
            }
            None => None,
        };

        let rule: Option<Handle<JsObject>> = obj.get_opt(cx, "iconRule")?;
        let rule = match rule {
            Some(rule) => Some(NodeXlsxFormatTypeRule::from_js_value(cx, rule)?),
            None => None,
        };

        Ok(Self {
            greater_than,
            no_icon,
            i_type,
            rule,
            i_type_index,
        })
    }
}

impl Into<ConditionalFormatCustomIcon> for CustomIcon {
    fn into(self) -> ConditionalFormatCustomIcon {
        let mut icon = ConditionalFormatCustomIcon::new();
        if let Some(i_type) = self.i_type {
            let index = self.i_type_index.unwrap();
            icon = icon.set_icon_type(i_type, index);
        }

        if self.greater_than {
            icon = icon.set_greater_than(true);
        }

        if self.no_icon {
            icon = icon.set_no_icon(true);
        }

        if let Some(rule) = self.rule {
            icon = icon.set_rule(rule.r_type, rule.value);
        }

        icon
    }
}

pub struct Icon {
    pub reverse: bool,
    pub show_icons_only: bool,
    icons: Vec<CustomIcon>,
    pub i_type: ConditionalFormatIconType,
    pub multi_range: Option<String>,
    pub stop_if_true: Option<bool>,
}

impl Icon {
    pub fn from_js_object(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let reverse: Handle<JsBoolean> = obj.get(cx, "reverse")?;
        let reverse = reverse.value(cx);

        let show_icons_only: Handle<JsBoolean> = obj.get(cx, "showIconsOnly")?;
        let show_icons_only = show_icons_only.value(cx);

        let icons: Handle<JsArray> = obj.get(cx, "icons")?;
        let icons = icons.to_vec(cx)?;
        let mut icons_vec = Vec::with_capacity(icons.len() as usize);
        for icon in icons {
            let icon = icon.downcast_or_throw::<JsObject, FunctionContext>(cx)?;
            let icon = CustomIcon::from_js_object(cx, icon)?;
            icons_vec.push(icon);
        }

        let i_type: Handle<JsString> = obj.get(cx, "iconType")?;
        let i_type = icon_from_js_string(cx, i_type)?;

        let multi_range: Option<Handle<JsString>> = obj.get_opt(cx, "multiRange")?;
        let multi_range = multi_range.map(|range| range.value(cx));

        let stop_if_true: Option<Handle<JsBoolean>> = obj.get_opt(cx, "stopIfTrue")?;
        let stop_if_true = stop_if_true.map(|stop| stop.value(cx));

        Ok(Self {
            reverse,
            show_icons_only,
            icons: icons_vec,
            i_type,
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
            let icon = Icon::from_js_object(cx, obj)?;
            let icon = NodeXlsxConditionalFormatType::IconSet(icon.into());
            c_format_map.insert(id, icon);
        }
        Ok(id)
    }
}

impl Into<ConditionalFormatIconSet> for Icon {
    fn into(self) -> ConditionalFormatIconSet {
        let mut icon_set = ConditionalFormatIconSet::new();
        icon_set = icon_set.set_icon_type(self.i_type);

        if self.reverse {
            icon_set = icon_set.reverse_icons(true);
        }

        if self.show_icons_only {
            icon_set = icon_set.show_icons_only(true);
        }

        let icons: Vec<ConditionalFormatCustomIcon> =
            self.icons.into_iter().map(|icon| icon.into()).collect();

        if icons.len() > 0 {
            icon_set = icon_set.set_icons(&icons);
        }

        if let Some(multi_range) = self.multi_range {
            icon_set = icon_set.set_multi_range(multi_range);
        }

        if let Some(stop_if_true) = self.stop_if_true {
            icon_set = icon_set.set_stop_if_true(stop_if_true);
        }

        icon_set
    }
}

fn icon_from_js_string(
    cx: &mut FunctionContext,
    value: Handle<JsString>,
) -> NeonResult<ConditionalFormatIconType> {
    let value = value.value(cx);
    match value.as_str() {
        "threeArrows" => Ok(ConditionalFormatIconType::ThreeArrows),
        "threeArrowsGray" => Ok(ConditionalFormatIconType::ThreeArrowsGray),
        "threeFlags" => Ok(ConditionalFormatIconType::ThreeFlags),
        "threeTrafficLights" => Ok(ConditionalFormatIconType::ThreeTrafficLights),
        "threeTrafficLightsWithRim" => Ok(ConditionalFormatIconType::ThreeTrafficLightsWithRim),
        "threeSigns" => Ok(ConditionalFormatIconType::ThreeSigns),
        "threeSymbolsCircled" => Ok(ConditionalFormatIconType::ThreeSymbolsCircled),
        "threeSymbols" => Ok(ConditionalFormatIconType::ThreeSymbols),
        "threeStars" => Ok(ConditionalFormatIconType::ThreeStars),
        "threeTriangles" => Ok(ConditionalFormatIconType::ThreeTriangles),
        "fourArrows" => Ok(ConditionalFormatIconType::FourArrows),
        "fourArrowsGray" => Ok(ConditionalFormatIconType::FourArrowsGray),
        "fourRedToBlack" => Ok(ConditionalFormatIconType::FourRedToBlack),
        "fourHistograms" => Ok(ConditionalFormatIconType::FourHistograms),
        "fourTrafficLights" => Ok(ConditionalFormatIconType::FourTrafficLights),
        "fiveArrows" => Ok(ConditionalFormatIconType::FiveArrows),
        "fiveArrowsGray" => Ok(ConditionalFormatIconType::FiveArrowsGray),
        "fiveHistograms" => Ok(ConditionalFormatIconType::FiveHistograms),
        "fiveQuadrants" => Ok(ConditionalFormatIconType::FiveQuadrants),
        "fiveBoxes" => Ok(ConditionalFormatIconType::FiveBoxes),
        _ => {
            let err = format!("Invalid ConditionalFormatIconType: {}", value);
            cx.throw_error(err)
        }
    }
}
