pub fn emu_to_px(emu: i32) -> i32 {
    let dpi = 96.0;
    (emu as f32 * (dpi / 914400.0)) as i32
}

pub fn docx_pt_to_px(v: i32) -> i32 {
    let pt = v / 2;
    pt_to_px(pt)
}

pub fn indent_to_px(v: i32) -> i32 {
    v / 15
}

fn pt_to_px(pt: i32) -> i32 {
    (pt as f32 / 0.75) as i32
}

pub fn set_panic_hook() {
    // When the `console_error_panic_hook` feature is enabled, we can call the
    // `set_panic_hook` function at least once during initialization, and then
    // we will get better error messages if our code ever panics.
    //
    // For more details see
    // https://github.com/rustwasm/console_error_panic_hook#readme
    #[cfg(feature = "console_error_panic_hook")]
    console_error_panic_hook::set_once();
}
