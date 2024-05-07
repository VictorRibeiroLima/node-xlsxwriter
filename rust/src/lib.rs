use neon::prelude::*;
use neon::types::buffer::TypedArray;
use node_xlsx::NodeXlsxWorkbook;

mod node_xlsx;

fn save_to_buffer(mut cx: FunctionContext) -> JsResult<JsPromise> {
    let js_obj: Handle<JsObject> = cx.argument(0)?;
    let workbook = NodeXlsxWorkbook::from_js_object(&mut cx, js_obj)?;

    let promise = cx
        .task(move || workbook.save_to_buffer())
        .promise(move |mut cx, greeting| match greeting {
            Ok(buffer) => {
                let mut js_buffer = cx.buffer(buffer.len())?;
                for (i, elem) in js_buffer.as_mut_slice(&mut cx).iter_mut().enumerate() {
                    *elem = buffer[i];
                }
                Ok(js_buffer)
            }
            Err(err) => cx.throw_error(err.to_string()),
        });

    Ok(promise)
}

fn save_to_buffer_sync(mut cx: FunctionContext) -> JsResult<JsBuffer> {
    let js_obj: Handle<JsObject> = cx.argument(0)?;
    let workbook = NodeXlsxWorkbook::from_js_object(&mut cx, js_obj)?;

    let buffer = workbook.save_to_buffer();

    let buffer = match buffer {
        Ok(buffer) => buffer,
        Err(err) => return cx.throw_error(err.to_string()),
    };

    let mut js_buffer = cx.buffer(buffer.len())?;
    for (i, elem) in js_buffer.as_mut_slice(&mut cx).iter_mut().enumerate() {
        *elem = buffer[i];
    }

    Ok(js_buffer)
}

fn save_to_file_sync(mut cx: FunctionContext) -> JsResult<JsUndefined> {
    let js_obj: Handle<JsObject> = cx.argument(0)?;
    let path: Handle<JsString> = cx.argument(1)?;
    let path = path.value(&mut cx);

    let workbook = NodeXlsxWorkbook::from_js_object(&mut cx, js_obj)?;

    match workbook.save_to_file(&path) {
        Ok(_) => Ok(cx.undefined()),
        Err(err) => return cx.throw_error(err.to_string()),
    }
}

fn save_to_file(mut cx: FunctionContext) -> JsResult<JsPromise> {
    let js_obj: Handle<JsObject> = cx.argument(0)?;
    let path: Handle<JsString> = cx.argument(1)?;
    let path = path.value(&mut cx);

    let workbook = NodeXlsxWorkbook::from_js_object(&mut cx, js_obj)?;

    let promise = cx
        .task(move || workbook.save_to_file(&path))
        .promise(|mut cx, result| match result {
            Ok(_) => Ok(cx.undefined()),
            Err(err) => cx.throw_error(err.to_string()),
        });

    Ok(promise)
}

fn save_to_bas64(mut cx: FunctionContext) -> JsResult<JsPromise> {
    let js_obj: Handle<JsObject> = cx.argument(0)?;
    let workbook = NodeXlsxWorkbook::from_js_object(&mut cx, js_obj)?;

    let promise = cx
        .task(move || workbook.save_to_base64())
        .promise(move |mut cx, greeting| match greeting {
            Ok(base64) => Ok(cx.string(base64)),
            Err(err) => cx.throw_error(err.to_string()),
        });

    Ok(promise)
}

fn save_to_bas64_sync(mut cx: FunctionContext) -> JsResult<JsString> {
    let js_obj: Handle<JsObject> = cx.argument(0)?;
    let workbook = NodeXlsxWorkbook::from_js_object(&mut cx, js_obj)?;

    let base64 = workbook.save_to_base64();

    let base64 = match base64 {
        Ok(base64) => base64,
        Err(err) => return cx.throw_error(err.to_string()),
    };

    Ok(cx.string(base64))
}

#[neon::main]
fn main(mut cx: ModuleContext) -> NeonResult<()> {
    cx.export_function("saveToBuffer", save_to_buffer)?;
    cx.export_function("saveToBufferSync", save_to_buffer_sync)?;
    cx.export_function("saveToFileSync", save_to_file_sync)?;
    cx.export_function("saveToFile", save_to_file)?;
    cx.export_function("saveToBase64", save_to_bas64)?;
    cx.export_function("saveToBase64Sync", save_to_bas64_sync)?;
    Ok(())
}
