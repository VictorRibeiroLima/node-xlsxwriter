use base64::{engine::general_purpose, Engine};
use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsArray, JsObject, JsValue},
};
use rust_xlsxwriter::Workbook;

use self::{error::NodeXlsxError, sheet::NodeXlsxSheet};

mod border;
mod cell;
mod color;
mod conditional_format;
mod error;
mod format;
mod sheet;
mod table;
mod types;
mod util;

pub struct NodeXlsxWorkbook {
    sheets: Vec<NodeXlsxSheet>,
}

impl NodeXlsxWorkbook {
    pub fn from_js_object<'b>(
        cx: &mut FunctionContext<'b>,
        obj: Handle<JsObject>,
    ) -> NeonResult<Self> {
        let mut inner_sheets = vec![];

        let sheets: Handle<JsArray> = obj.get(cx, "sheets")?;
        let sheets: Vec<Handle<JsValue>> = sheets.to_vec(cx)?;
        for sheet in sheets {
            let sheet = sheet.downcast_or_throw::<JsObject, FunctionContext>(cx)?;
            let sheet = NodeXlsxSheet::from_js_object(cx, sheet)?;
            inner_sheets.push(sheet);
        }
        return Ok(Self {
            sheets: inner_sheets,
        });
    }

    pub fn save_to_buffer(self) -> Result<Vec<u8>, NodeXlsxError> {
        let mut workbook = self.parse()?;
        let buffer = workbook.save_to_buffer()?;
        return Ok(buffer);
    }

    pub fn save_to_file(self, path: &str) -> Result<(), NodeXlsxError> {
        let mut workbook = self.parse()?;
        workbook.save(path)?;
        return Ok(());
    }

    pub fn save_to_base64(self) -> Result<String, NodeXlsxError> {
        let buffer = self.save_to_buffer()?;
        let base64 = general_purpose::STANDARD.encode(buffer);
        return Ok(base64);
    }

    fn parse(self) -> Result<Workbook, NodeXlsxError> {
        let mut workbook = rust_xlsxwriter::Workbook::new();
        for sheet in self.sheets {
            let worksheet = sheet.try_into().unwrap(); //TODO: Handle error
            workbook.push_worksheet(worksheet);
        }
        return Ok(workbook);
    }
}
