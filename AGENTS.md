# AGENTS.md — ISEL-to-Gcode

ISEL/CAM programlarını ISIN'nin ısıtıcı sarma makinelerinin okuduğu G-code'a çeviren proje.
Durum, açık işler, bug geçmişi ve test sonuçları burada değil Obsidian notunda tutulur:
`Projeler/ISEL_to_Gcode.md`. Burası yalnız kural kitabı.

## Nerede ne var
- `rust/` programın tamamı: kütüphane + tek çalıştırılabilir `ISEL_to_Gcode`.
  `src/convert.rs` dönüşüm, `src/gcode.rs` önizleme için G-code ayrıştırıcısı,
  `src/preview.rs` çizim sahnesi (pencere olmadan test edilebilir), `src/app.rs` eframe arayüzü ve
  CLI girişi, `src/main.rs` ince giriş noktası.
- `tests/fixtures/` bayt golden'ları (sentetik fixture + iki gerçek müşteri programı). Sözleşme bunlar;
  `rust/tests/` bunları çalıştırır.
- Golden'ları üreten Python programı repoda değil. Yorumlardaki "python line N" referansları o dosyaya
  ait; kurtarma: `git show 9b1fa1d^:isel_to_gcode_gui.py`.

## Komutlar
- Yerel derleme/test: `export PATH="$HOME/tools/w64devkit/bin:$PATH"; cd rust && cargo test && cargo build --release`
- Windows GNU yalnız. `rust/.cargo/config.toml` içindeki `-C link-self-contained=yes` korunur.
- CLI: `ISEL_to_Gcode.exe convert <in.isel> <out.ngc> [--arcs]` — bayraklar iki yoldan sonra gelir.
- Release: `rust/Cargo.toml` içindeki `version` tek sürüm kaynağı; commit → `git tag -a vX.Y.Z` → etiketi
  push → `gh release create vX.Y.Z`. Yalnızca *yayınlanan* release CI derlemesini tetikler ve
  `ISEL_to_Gcode.exe`'i release'e ekler.

## Yap
- Çıktıyı `tests/fixtures/*expected*` ile bayt bayt aynı tut.
- Referans programın tuhaflıklarını koru (örneğin negatif sıfırın `-0.000` yazılması); referansa uymak
  güzelliğe uymaktan önce gelir.
- Dönüşümü işçi iş parçacığında yap, arayüz iş parçacığında değil; önizleme işçinin sonucunu kullanır.
- `eframe::run_native` dönüşünü kontrol et; yok sayılan hata penceresiz sessiz çıkış demektir.
- Başlangıç günlüğünü (`%TEMP%\ISEL_to_Gcode_startup.log`), panic hook'unu ve hata kutusunu koru;
  pencere açılmayan bir makinede tek kanıt bunlar.
- Render sırası Glow (OpenGL) önce, wgpu yedek; `ISEL_RENDERER=glow|wgpu` ile geçici olarak değiştirilir.
- Renkleri tek yerden ver: `src/theme.rs`. Widget'ları elle boyama; `theme::apply()` hem Dark hem Light
  kovasına yazar ve tema tercihini sabitler — Windows light modunda egui'nin beyaz varsayılanları (progress
  yolu, checkbox) aksi halde geri sızıyor.
- Penceresiz makinede arayüz kanıtı: `ISEL_SHOT=yol.bmp` değişkeniyle uygulama tek kare ekran görüntüsü
  yazıp çıkar. BMP yazıcısının kanal/satır sırasını `app::tests::shot_bmp_is_bottom_up_bgr` korur.
- Exe dosya ikonu `rust/build.rs`'in `windres` ile gömdüğü Win32 kaynağından gelir (`assets/icon.ico`);
  `with_icon` yalnız canlı pencereyi boyar. `windres` PATH'te olmalı (yerelde w64devkit, CI'da `C:\mingw64`) —
  yoksa build.rs bilerek düşer, ikonsuz exe sessizce yayınlanmaz.
- Önizleme penceresi ayrı bir OS penceresidir (`show_viewport_immediate`); `egui::Window`'a geri döndürme,
  pencere ana pencerenin içine kıstırılır. Başsız kanıt: `ISEL_PREVIEW=<out.ngc>` önizlemeyi açılışta açar,
  `EnumWindows` ile ikinci üst-düzey pencere görülür.
- `.gitattributes` içindeki `.ngc`/`.isel`/`.ncp` CRLF pinlerini koru.
- Kod değişikliğinden sonra codegraph'ı `sync` ile tazele.

## Yapma
- Python/.NET/Electron katmanını geri getirme; teslim tek dosya exe, kurulum istemiyor.
- Test geçsin diye golden'ı güncel Rust çıktısıyla üretme.
- Yerel ya da CI derlemesini MSVC `link.exe`'ine geçirme; bu makinede çöküyor.
- Sentetik fixture golden'larını (`sample_expected.*`) silme; gerçek programların girmediği yolları
  (VEL yokken varsayılan ilerleme, R yayları, sıfır uzunluk koruması) onlar kapsar.
- `graphify-out/`, `graph_update.log`, `.codegraph/`, `rust/target/` commit edilmez.
- `cargo fmt` çalıştırma: kod elle biçimli, rustfmt tüm ağacı yeniden yazar ve gerçek değişikliği gömer.
- Bu projede subagent kullanma.

## Kod Grafiği (graphify) — OTOMATİK KULLANIM

- **Proje dosyası açmadan ÖNCE** mimariye hızlı bak (cron ile her saat güncellenir):
  ```bash
  read_file graphify-out/GRAPH_REPORT.md   # God Nodes + Communities (ilk 30 satır yeterli)
  ```
- **Büyük refactor / çok dosyalı değişiklikten ÖNCE** grafiği sorgula:
  ```bash
  C:/Users/Ai/AppData/Local/hermes/venvs/graphify/Scripts/python.exe -m graphify affected "modul_veya_fonksiyon" --depth 2 --graph graphify-out/graph.json
  C:/Users/Ai/AppData/Local/hermes/venvs/graphify/Scripts/python.exe -m graphify query "soru" --graph graphify-out/graph.json
  ```
- Küçük tek dosya değişikliklerinde gerek yok — `search_files` yeterli.
