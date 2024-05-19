use std::collections::HashMap;

use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsArray, JsBoolean, JsNumber, JsObject, JsString, JsValue},
};

use crate::node_xlsx::util::create_format;

pub enum SizeType {
    AUTO,
    PX,
}

#[derive(Debug, Clone, Copy)]
pub enum Type {
    ROW,
    COLUMN,
}

pub struct SizeConfig {
    pub value: f64,
    pub unit: SizeType,
}

impl SizeConfig {
    pub fn from_js_object(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let value: Handle<JsNumber> = obj.get(cx, "value")?;
        let value = value.value(cx);

        let size_type: Option<Handle<JsString>> = obj.get_opt(cx, "unit")?;
        let unit = match size_type {
            Some(size_type) => {
                let size_type = size_type.value(cx).to_lowercase();
                match size_type.as_str() {
                    "auto" => SizeType::AUTO,
                    "px" => SizeType::PX,
                    _ => {
                        let error_message = format!("Invalid sizeType: {}", size_type);
                        return cx.throw_error(error_message);
                    }
                }
            }
            None => SizeType::AUTO,
        };

        Ok(Self { value, unit })
    }
}

pub struct RowColumnConfig {
    pub index: u32,
    pub config_type: Type,
    pub size: Option<SizeConfig>,
    pub format: Option<u32>,
    pub hidden: Option<bool>,
}

impl RowColumnConfig {
    pub fn from_js_array(
        cx: &mut FunctionContext,
        arr: Handle<JsArray>,
        config_type: Type,
        format_map: &mut HashMap<u32, rust_xlsxwriter::Format>,
    ) -> NeonResult<Vec<Self>> {
        let mut row_column_configs = vec![];

        let js_array: Vec<Handle<JsValue>> = arr.to_vec(cx)?;
        for js_value in js_array {
            let js_value = js_value.downcast_or_throw::<JsObject, FunctionContext>(cx)?;
            let row_column_config =
                RowColumnConfig::from_js_object(cx, js_value, config_type, format_map)?;
            row_column_configs.push(row_column_config);
        }

        Ok(row_column_configs)
    }

    pub fn from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        config_type: Type,
        format_map: &mut HashMap<u32, rust_xlsxwriter::Format>,
    ) -> NeonResult<Self> {
        let result = Self::inner_from_js_object(cx, obj, config_type, format_map);
        match result {
            Ok(config) => Ok(config),
            Err(error) => {
                let error = format!(
                    "Error parsing RowColumnConfig: {:?} with error:\n  {}",
                    obj, error
                );
                let js_string = cx.string(error);
                cx.throw(js_string)
            }
        }
    }

    fn inner_from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        config_type: Type,
        format_map: &mut HashMap<u32, rust_xlsxwriter::Format>,
    ) -> NeonResult<Self> {
        let index: Handle<JsNumber> = obj.get(cx, "index")?;
        let index = index.value(cx) as u32;

        let size: Option<Handle<JsObject>> = obj.get_opt(cx, "size")?;
        let size = match size {
            Some(size) => Some(SizeConfig::from_js_object(cx, size)?),
            None => None,
        };

        let format: Option<Handle<JsObject>> = obj.get_opt(cx, "format")?;
        let format = match format {
            Some(format) => Some(create_format(cx, format, format_map)?),
            None => None,
        };

        let hidden: Option<Handle<JsBoolean>> = obj.get_opt(cx, "hidden")?;
        let hidden = match hidden {
            Some(hidden) => Some(hidden.value(cx)),
            None => None,
        };

        Ok(Self {
            index,
            size,
            format,
            hidden,
            config_type,
        })
    }

    pub fn write_to_sheet(
        self,
        sheet: &mut rust_xlsxwriter::Worksheet,
        format_map: &HashMap<u32, rust_xlsxwriter::Format>,
    ) -> Result<(), rust_xlsxwriter::XlsxError> {
        match self.config_type {
            Type::ROW => {
                if let Some(size) = self.size {
                    match size.unit {
                        SizeType::AUTO => sheet.set_row_height(self.index, size.value)?,
                        SizeType::PX => {
                            sheet.set_row_height_pixels(self.index, size.value as u16)?
                        }
                    };
                }

                if let Some(format) = self.format {
                    let format = format_map.get(&format).unwrap();
                    sheet.set_row_format(self.index, format)?;
                }

                if let Some(hidden) = self.hidden {
                    if hidden {
                        sheet.set_row_hidden(self.index)?;
                    }
                }
            }
            Type::COLUMN => {
                if let Some(size) = self.size {
                    match size.unit {
                        SizeType::AUTO => sheet.set_column_width(self.index as u16, size.value)?,
                        SizeType::PX => {
                            sheet.set_column_width_pixels(self.index as u16, size.value as u16)?
                        }
                    };
                }

                if let Some(format) = self.format {
                    let format = format_map.get(&format).unwrap();
                    sheet.set_column_format(self.index as u16, format)?;
                }

                if let Some(hidden) = self.hidden {
                    if hidden {
                        sheet.set_column_hidden(self.index as u16)?;
                    }
                }
            }
        };

        Ok(())
    }
}
