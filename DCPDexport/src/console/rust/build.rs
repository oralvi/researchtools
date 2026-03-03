fn main() {
    #[cfg(windows)]
    {
        if std::path::Path::new("../../DCPDexport/icon.ico").exists() {
            let mut res = winres::WindowsResource::new();
            res.set_icon("../../DCPDexport/icon.ico");
            res.compile().unwrap();
        }
    }
}
