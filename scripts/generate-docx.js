/**
 * Tạo báo cáo Word (.docx) theo yêu cầu môn HTRQD:
 * data abstraction, task abstraction, idiom, đánh giá + link dashboard.
 */
const fs = require("fs");
const path = require("path");
const {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  Footer,
  Header,
  HeadingLevel,
  LevelFormat,
  Packer,
  PageNumber,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
  convertInchesToTwip,
} = require("docx");

const OUT = path.join(__dirname, "..", "BaoCao_PhanTichThietKeTrucQuan.docx");
const DASHBOARD = "http://localhost:5173";

const BLUE = "0B4F6C";
const NAVY = "1B2A4A";
const HEAD_BG = "E8EEF2";
const THIN = { style: BorderStyle.SINGLE, size: 8, color: "444444" };
const BORDERS = { top: THIN, bottom: THIN, left: THIN, right: THIN };
const MARGINS = { top: 60, bottom: 60, left: 80, right: 80 };

const run = (text, opt = {}) =>
  new TextRun({
    text,
    font: "Times New Roman",
    size: opt.size || 26,
    bold: !!opt.bold,
    italics: !!opt.italics,
    color: opt.color,
    underline: opt.underline ? {} : undefined,
  });

const p = (children, opt = {}) =>
  new Paragraph({
    spacing: { after: opt.after ?? 160, before: opt.before ?? 0, line: 360 },
    alignment: opt.align,
    indent: opt.indent,
    children: Array.isArray(children) ? children : [run(children)],
  });

const h1 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 280, after: 140, line: 360 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "C9A227", space: 4 } },
    children: [run(text, { bold: true, size: 28, color: NAVY })],
  });

const h2 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 220, after: 100, line: 360 },
    children: [run(text, { bold: true, size: 26, color: BLUE })],
  });

const h3 = (text) =>
  new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 160, after: 80, line: 360 },
    children: [run(text, { bold: true, italics: true, size: 26 })],
  });

function cell(text, opt = {}) {
  const content = Array.isArray(text) ? text : String(text);
  return new TableCell({
    borders: BORDERS,
    width: { size: opt.width || 2500, type: WidthType.DXA },
    shading: opt.header ? { type: ShadingType.CLEAR, fill: HEAD_BG } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: MARGINS,
    columnSpan: opt.span,
    children: [
      new Paragraph({
        spacing: { after: 40, line: 276 },
        children: Array.isArray(content)
          ? content
          : [run(content, { bold: !!opt.header || !!opt.bold, size: opt.header ? 22 : 22 })],
      }),
    ],
  });
}

function table(headers, rows, widths) {
  const w = widths || headers.map(() => Math.floor(9360 / headers.length));
  const head = new TableRow({
    children: headers.map((h, i) => cell(h, { header: true, width: w[i] })),
  });
  const body = rows.map(
    (r) =>
      new TableRow({
        children: r.map((c, i) => cell(c, { width: w[i], bold: i === 0 && !!r._boldFirst })),
      })
  );
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: w,
    rows: [head, ...body],
  });
}

const bullet = (parts) =>
  new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80, line: 360 },
    children: Array.isArray(parts) ? parts : [run(parts)],
  });

const numbered = (parts, ref = "nums") =>
  new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 80, line: 360 },
    children: Array.isArray(parts) ? parts : [run(parts)],
  });

const spacer = () => new Paragraph({ spacing: { after: 80 }, children: [] });

const doc = new Document({
  creator: "Nhóm HTRQD — 25C12012, 25C12038, 25C12039",
  title: "Báo cáo phân tích và thiết kế trực quan — Atlas khí tượng Việt Nam",
  description:
    "Báo cáo gồm data abstraction, task abstraction, idiom và đánh giá theo mô hình Munzner.",
  styles: {
    default: {
      document: {
        styles: [
          {
            id: "Normal",
            run: { font: "Times New Roman", size: 26 },
            paragraph: { spacing: { line: 360 } },
          },
        ],
      },
    },
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 420, hanging: 220 } } },
          },
        ],
      },
      {
        reference: "nums",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 420, hanging: 220 } } },
          },
        ],
      },
      {
        reference: "refs",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "[%1]",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 480, hanging: 280 } } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: convertInchesToTwip(8.27), height: convertInchesToTwip(11.69) },
          margin: {
            top: convertInchesToTwip(0.9),
            bottom: convertInchesToTwip(0.9),
            left: convertInchesToTwip(1.1),
            right: convertInchesToTwip(0.9),
          },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [run("HTRQD · Báo cáo phân tích và thiết kế trực quan", { size: 18, italics: true, color: "666666" })],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                run("Trang ", { size: 18, color: "666666" }),
                new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 18, color: "666666" }),
              ],
            }),
          ],
        }),
      },
      children: [
        p("BỘ GIÁO DỤC VÀ ĐÀO TẠO", { align: AlignmentType.CENTER, after: 40 }),
        p("CHƯƠNG TRÌNH ĐÀO TẠO SAU ĐẠI HỌC", { align: AlignmentType.CENTER, after: 40 }),
        p("Môn học: Hình thức / Trực quan hóa dữ liệu (HTRQD)", { align: AlignmentType.CENTER, after: 280 }),
        p([run("BÁO CÁO", { bold: true, size: 36, color: NAVY })], { align: AlignmentType.CENTER, after: 80 }),
        p([run("PHÂN TÍCH VÀ THIẾT KẾ TRỰC QUAN", { bold: true, size: 32, color: NAVY })], {
          align: AlignmentType.CENTER,
          after: 160,
        }),
        p([run("Dashboard khí tượng 63 tỉnh Việt Nam (21/04/2024 – 04/06/2025)", { italics: true, size: 26 })], {
          align: AlignmentType.CENTER,
          after: 80,
        }),
        p(
          [
            run(
              "Áp dụng mô hình lồng nhau của Munzner: domain situation · data/task abstraction · idiom · evaluation",
              { size: 22, color: "444444" }
            ),
          ],
          { align: AlignmentType.CENTER, after: 280 }
        ),

        table(
          ["Mã học viên", "Họ và tên", "Task (slide lý thuyết)", "Idiom chính"],
          [
            ["25C12012", "Nguyễn Quốc Trạng", "Discover distribution", "Proportional symbol map + histogram"],
            ["25C12038", "Trương Thị Ngọc Viên", "Compare trends", "Aligned line chart + heatmap"],
            ["25C12039", "Phạm Quốc Vương", "Locate outliers", "Scatterplot + boxplot (IQR)"],
          ],
          [1600, 2600, 2500, 2660]
        ),
        spacer(),

        p(
          [
            run("Link dashboard: ", { bold: true }),
            new ExternalHyperlink({
              children: [run(DASHBOARD, { color: BLUE, underline: true })],
              link: DASHBOARD,
            }),
            run("  — chạy local bằng lệnh "),
            run("node serve.js", { italics: true }),
            run(" trong thư mục dashboard (chi tiết mục 8). Sau khi triển khai GitHub Pages/Netlify, thay URL này trước khi nộp."),
          ],
          { after: 80 }
        ),
        p(
          [
            run("Yêu cầu đề bài đã đáp ứng: ", { bold: true }),
            run(
              "mỗi thành viên chọn 1 task trên slide cuối lý thuyết (cặp {action, target}), thiết kế idiom – cài đặt – đánh giá; ba task được tích hợp một dashboard có tương tác liên kết (điểm bonus); báo cáo gồm data abstraction, task abstraction, idiom và đánh giá, kèm link dashboard."
            ),
          ],
          { after: 200 }
        ),

        h1("1. Tình huống miền (Domain situation)"),
        p(
          "Người dùng mục tiêu gồm nhà phân tích khí tượng / quy hoạch lãnh thổ và học viên môn trực quan hóa. Họ cần khám phá (discover) chứ không chỉ trình bày số liệu đã biết (present). Các câu hỏi miền điển hình: tỉnh nào nóng hoặc mưa hơn; mùa mưa Bắc–Nam lệch pha như thế nào; ngày nào là cực đoan. Dữ liệu đầu vào là bảng quan sát ngày tại 63 tỉnh Việt Nam, nguồn WeatherAPI, file df_weather_fixed_utf8.csv đã chuẩn hóa UTF-8."
        ),
        p(
          "Theo nested model, sai ở tầng domain (hiểu sai nhu cầu người dùng) sẽ lan xuống mọi tầng dưới. Nhóm neo domain vào ba nhu cầu phân tích khí hậu: phân bố không gian, xu hướng theo thời gian, và sự kiện lệch. Ba nhu cầu này ánh xạ trực tiếp sang ba cặp {action, target} trên slide cuối chương Task Abstraction."
        ),

        h1("2. Data abstraction (What)"),
        p(
          "Munzner yêu cầu tách “dữ liệu thô của miền” khỏi “cái được đưa vào visualization”. Dataset gốc là table: mỗi dòng là một item (tỉnh × ngày), mỗi cột là một attribute. Đây không phải network hay field liên tục; không gian được gắn vào item qua hai thuộc tính lat/lon chứ không phải lưới raster."
        ),

        h2("2.1. Items và thuộc tính gốc"),
        table(
          ["Thành phần", "Giá trị quan sát", "Kiểu trừu tượng (Munzner)"],
          [
            ["Số item", "26.018 dòng (63 tỉnh × 409 ngày, thiếu 1 ô)", "table items"],
            ["Thời gian (date)", "21/04/2024 → 04/06/2025", "ordered sequential (key)"],
            ["location.name", "63 tỉnh/thành", "categorical key"],
            ["location.region", "6 vùng kinh tế–địa lý", "categorical"],
            ["location.terrain", "đồng bằng, ven biển, miền núi", "categorical"],
            ["lat, lon", "tọa độ trung tâm tỉnh", "quantitative spatial"],
            [
              "avgtemp / maxtemp / mintemp (°C)",
              "Tavg 25,1°C; Tmax 29,2°C; cực trị 43,4°C / −0,2°C",
              "quantitative sequential",
            ],
            [
              "totalprecip_mm",
              "TB 6,65 mm/ngày; max 285,2 mm (Quảng Trị, 27/10/2024)",
              "quantitative sequential, lệch phải",
            ],
            ["avghumidity, maxwind_kph, uv", "Độ ẩm TB 76,8%", "quantitative sequential"],
            ["condition.text", "21 nhãn thời tiết", "categorical"],
          ],
          [2600, 4000, 2760]
        ),
        spacer(),
        p(
          "Sáu vùng: Đồng bằng sông Cửu Long, Đông Nam Bộ, Đồng bằng sông Hồng, Bắc Trung Bộ và Duyên hải miền Trung, Tây Nguyên, Trung du và miền núi Bắc Bộ. Ba địa hình: đồng bằng, ven biển, miền núi. Availability: dataset tĩnh (static), không streaming."
        ),

        h2("2.2. Thuộc tính suy ra (derived)"),
        p(
          [
            run("Munzner nhấn mạnh: “don’t just use what you’re given”. "),
            run("Dashboard derive thêm các attribute phục vụ đúng task, không chỉ vẽ cột thô:"),
          ]
        ),
        bullet([run("Chỉ số tỉnh–ngày", { bold: true }), run(": mảng phẳng 63×409 cho tmax, tmin, tavg, precip, humidity, wind — cho phép lọc tương tác không quét lại CSV.")]),
        bullet([run("Trung bình theo tỉnh trong khoảng ngày", { bold: true }), run(" (item aggregation) — item của bản đồ.")]),
        bullet([run("Chuỗi ngày theo vùng", { bold: true }), run(" và "), run("trung bình trượt 7 ngày", { bold: true }), run(" — giảm nhiễu ngày, phục vụ compare trends.")]),
        bullet([run("Ma trận vùng × tháng", { bold: true }), run(" — aggregation hai chiều cho heatmap (summarize toàn kỳ).")]),
        bullet([
          run("Cờ outlier", { bold: true }),
          run(": mưa > Q3 + 1,5×IQR; nhiệt |z-score| ≥ 2,5 tính trên phân bố đang lọc. Không hard-code ngưỡng miền (ví dụ “mưa > 100 mm”) để idiom gắn với task locate chứ không với domain jargon."),
        ]),
        p(
          "Geometry tỉnh không có sẵn dạng polygon trong file gốc, nên vị trí không gian được trừu tượng hóa thành point geometry (lat/lon), không phải area geometry. Đây là lựa chọn abstraction có chủ đích: tránh bias diện tích của choropleth khi chưa có ranh giới sạch, và được đánh giá lại ở mục 6.",
          { after: 160 }
        ),

        h1("3. Task abstraction (Why)"),
        p(
          "Slide cuối chương Task Abstraction liệt kê các cặp {action, target} gốc: discover distribution · compare trends · locate outliers · browse topology. Nhóm chọn ba cặp đầu vì dataset là table + spatial attributes. Cặp browse topology dành cho network/tree nên loại, tránh “show the wrong thing”."
        ),
        p(
          "Mỗi task được viết lại theo ba tầng action độc lập (Analyze / Search / Query) rồi gắn target. Domain jargon được gỡ bỏ có hệ thống đúng rule of thumb của Munzner."
        ),

        h2("3.1. Task 1 — Nguyễn Quốc Trạng (25C12012) — Discover distribution"),
        p(
          [
            run("Câu hỏi miền: ", { bold: true }),
            run("Tỉnh nào nóng/mát, mưa nhiều/ít? Phân bố đó có phụ thuộc địa hình không?"),
          ]
        ),
        table(
          ["Tầng action", "Miền (domain)", "Trừu tượng (abstract)"],
          [
            ["Analyze", "Khám phá nóng–mát, mưa nhiều–ít trên cả nước", "Consume → Discover (chưa biết đáp án trước)"],
            ["Search", "Không khóa sẵn một tỉnh", "Explore / Browse (location unknown)"],
            ["Query", "Nhìn toàn quốc rồi đọc một tỉnh", "Summarize (all) rồi Identify (one)"],
            ["Target", "Nóng–lạnh và mưa theo không gian, địa hình", "Distribution của quantitative attribute trên spatial field"],
          ],
          [1800, 3400, 4160]
        ),
        spacer(),
        p(
          "Kết quả abstraction: {discover, distribution} kết hợp {summarize → identify}. Idiom phải cho overview không gian và vẫn đọc được một item (tỉnh)."
        ),

        h2("3.2. Task 2 — Trương Thị Ngọc Viên (25C12038) — Compare trends"),
        p(
          [
            run("Câu hỏi miền: ", { bold: true }),
            run("Mùa mưa miền Nam và rét mùa đông miền Bắc lệch pha như thế nào theo tháng?"),
          ]
        ),
        table(
          ["Tầng action", "Miền (domain)", "Trừu tượng (abstract)"],
          [
            ["Analyze", "Đối chiếu pha mùa giữa các vùng", "Discover; hỗ trợ Present khi giảng"],
            ["Search", "Đã biết 6 vùng", "Lookup / Browse (target known: vùng)"],
            ["Query", "Đối chiếu nhiều vùng, nhiều tháng", "Compare (some) và Summarize (heatmap toàn kỳ)"],
            ["Target", "Pha mùa, đỉnh mưa, hẫng nhiệt", "Trends của attribute theo ordered time key"],
          ],
          [1800, 3400, 4160]
        ),
        spacer(),
        p("Kết quả abstraction: {compare, trends}. Idiom phải dùng common-scale position theo thời gian, không dùng pie theo tháng (phá ordered key)."),

        h2("3.3. Task 3 — Phạm Quốc Vương (25C12039) — Locate outliers"),
        p(
          [
            run("Câu hỏi miền: ", { bold: true }),
            run("Ngày nào / tỉnh nào mưa cực đoan, nắng nóng hoặc rét đậm?"),
          ]
        ),
        table(
          ["Tầng action", "Miền (domain)", "Trừu tượng (abstract)"],
          [
            ["Analyze", "Tìm sự kiện thời tiết cực đoan", "Discover"],
            ["Search", "Không biết trước tỉnh–ngày nào", "Locate (đặc trưng đã biết: lệch phân bố; vị trí chưa biết)"],
            ["Query", "Đọc đúng một sự kiện rồi so vài sự kiện", "Identify rồi Compare"],
            ["Target", "Đợt lũ, nắng nóng, rét", "Outliers trên hai attribute (temp, precip)"],
          ],
          [1800, 3400, 4160]
        ),
        spacer(),
        p(
          "Ba task độc lập về abstraction nhưng chia sẻ dataset và bộ lọc, nên có thể ghép thành coordinated multiple views mà không đổi data abstraction."
        ),

        h1("4. Idiom (How)"),
        p(
          "Idiom gồm visual encoding (how to draw) và interaction (how to manipulate). Kênh mã hóa được chọn theo thứ tự hiệu lực Munzner/Mackinlay đối với dữ liệu định lượng: position trên trục chung ≫ length/size ≫ color tuần tự ≫ shape. Expressiveness: không dùng pie cho spatial field hay time key."
        ),

        h2("4.1. Task 1 — Bivariate proportional symbol map + histogram"),
        table(
          ["Thành phần idiom", "Lựa chọn", "Lý do (effectiveness / expressiveness)"],
          [
            ["Mark", "Point mark tại lat/lon", "Spatial data → spatial position; tránh bias diện tích choropleth"],
            ["Kênh màu", "Sequential RdYlBu / Teal = thuộc tính đang chọn", "Color ordered phù hợp quantitative sequential"],
            ["Kênh size", "Diện tích marker = mưa TB", "Mã hóa thuộc tính thứ hai; size yếu hơn position nên chỉ phụ"],
            ["Kênh shape", "Hình = địa hình (tròn / thoi / vuông)", "Categorical, số hạng mục = 3 nên shape còn phân biệt được"],
            ["Idiom phụ", "Histogram tần suất tỉnh–ngày", "Summarize distribution; bản đồ không đếm tần suất tốt"],
            ["Tương tác", "Click tỉnh → filter toàn dashboard", "Select + connect (linked views)"],
          ],
          [2200, 3300, 3860]
        ),
        spacer(),
        p(
          "Phương án loại: pie (không mã hóa không gian); choropleth (thiếu polygon sạch, diện tích tỉnh bị đọc như magnitude). Zoom bản đồ geo bị tắt có chủ đích: zoomInGeo kết hợp trục lat/lon khóa làm “văng” lãnh thổ Việt Nam và không zoom out được — đó là threat ở tầng algorithm/idiom, đã vá để task discover distribution không bị gãy."
        ),

        h2("4.2. Task 2 — Aligned line chart + heatmap"),
        table(
          ["Thành phần idiom", "Lựa chọn", "Lý do"],
          [
            ["Mark", "Line trên trục thời gian chung", "Position on common scale — kênh mạnh nhất để compare trends"],
            ["Derived", "Trung bình trượt 7 ngày", "Tách signal khỏi noise ngày"],
            ["Idiom phụ", "Heatmap vùng × tháng", "Matrix alignment; color = magnitude; tốt cho summarize toàn kỳ"],
            ["Tương tác", "Lọc vùng/ngày; overlay nét đứt khi khóa tỉnh", "Superimpose + filter"],
          ],
          [2200, 3300, 3860]
        ),
        spacer(),
        p(
          "Phương án loại: stacked area (trend của lớp trên bị cộng dồn, dễ đọc sai); pie theo tháng (phá ordered time). Heatmap kém line ở đọc slope nên chỉ đóng vai trò tóm tắt, không thay line."
        ),

        h2("4.3. Task 3 — Scatterplot + boxplot + bảng xếp hạng"),
        table(
          ["Thành phần idiom", "Lựa chọn", "Lý do"],
          [
            ["Scatter", "x = Tavg, y = precip", "Hai kênh position — thấy correlation và điểm lệch"],
            ["Màu scatter", "Nhị phân inlier / outlier", "Target là outliers, không phải scale liên tục"],
            ["Boxplot", "Mưa theo địa hình, hiện fences IQR", "Idiom cổ điển cho distribution + explicit outlier"],
            ["Bảng 15 dòng", "Tên tỉnh, ngày, lý do lệch", "Identify bằng identity channel; vis không thay lookup giá trị"],
            ["Tương tác", "Lasso/box trên scatter → nhấn tỉnh trên map", "Brushing & linking"],
          ],
          [2200, 3300, 3860]
        ),
        spacer(),
        p("Phương án loại: chỉ sort bảng (không thấy tương quan hai biến); chỉ line (giấu điểm lệch đa biến)."),

        h2("4.4. Idiom tương tác tích hợp (điểm bonus)"),
        p("Ba task không đứng riêng mà dùng coordinated multiple views trên một dashboard:"),
        bullet("Shared filter: khoảng ngày, vùng, địa hình, metric màu — mọi view đọc cùng data abstraction đã lọc."),
        bullet("Select trên bản đồ → filter scatter, histogram, trend (details-on-demand + connect)."),
        bullet("Brush trên scatter → highlight tỉnh trên bản đồ."),
        bullet("Khóa tỉnh: overlay đường nét đứt trên trend để compare một item với vùng."),
        p(
          "Đây là interaction idiom của Shneiderman (overview first, zoom and filter, then details-on-demand) đặt lên nested model: overview = KPI + map + heatmap; filter = ngày/vùng/địa hình; details = bảng outlier và hover."
        ),

        h1("5. Algorithm (tầng trong cùng)"),
        p("Tầng algorithm không phải trọng tâm báo cáo nhưng cần ghi để hoàn nested model và nêu threat tương ứng."),
        bullet("Tiền xử lý Node.js: CSV UTF-8 → JSON đặc (mảng số), heatmap và region series; strip BOM."),
        bullet("Sửa nhãn vùng lỗi “Tr [*]ung du và miền núi Bắc Bộ”."),
        bullet("Vẽ client-side Plotly; scatter dùng WebGL (scattergl) vì ~20.000 điểm."),
        bullet("Outlier được tính lại theo bộ lọc hiện tại (dynamic derive)."),
        bullet("JSON ~0,9 MB; dashboard tĩnh, chạy bằng node serve.js, không cần backend."),

        h1("6. Đánh giá (Evaluation)"),
        p(
          "Nested model đòi hỏi validate đúng tầng. Không dùng thời gian chạy hệ thống để chứng minh idiom tốt, và không dùng user study phòng thí nghiệm để chứng minh abstraction đúng (mismatch đã được Munzner cảnh báo)."
        ),

        h2("6.1. Tầng domain — face validity"),
        p(
          "Ba câu hỏi (phân bố không gian, pha mùa, cực đoan) khớp người dùng khí tượng/học viên vis. Threat: nhóm không phỏng vấn dự báo viên thật (misunderstood needs). Giảm threat bằng cách neo task vào taxonomy có sẵn trên slide lý thuyết thay vì bịa task ad hoc."
        ),

        h2("6.2. Tầng abstraction — đối chiếu pattern miền"),
        p("Vis “đúng cái cần hiện” nếu tái tạo được các pattern khí hậu Việt Nam đã biết — đây là validation qualitative sau khi cài đặt:"),
        bullet("Nam Bộ / ĐBSCL ấm hơn (~27,6–27,8°C) so với Trung du miền núi Bắc Bộ (~22,3°C) và Tây Nguyên (~22,8°C)."),
        bullet("Miền núi mát hơn đồng bằng khoảng 4°C (22,7 so với 26,5) — distribution theo terrain."),
        bullet("Nắng nóng sớm: Hòa Bình 43,4°C ngày 28/04/2024."),
        bullet("Rét: Hà Giang −0,2°C ngày 12/01/2025."),
        bullet("Mưa cực đoan: Quảng Trị 285,2 mm ngày 27/10/2024 — đúng cửa sổ mưa lũ miền Trung."),
        p(
          "Các pattern xuất hiện trên map/histogram (task 1), lệch pha đường/heatmap (task 2), và điểm đỏ + bảng (task 3). Abstraction không rơi vào lỗi “showing the wrong thing”."
        ),

        h2("6.3. Tầng idiom — so với phương án khác"),
        table(
          ["Task", "Phương án đã loại", "Threat nếu chọn phương án đó"],
          [
            ["Discover distribution", "Pie; choropleth diện tích", "Pie không express spatial field; choropleth bias diện tích"],
            ["Compare trends", "Stacked area; pie theo tháng", "Stack làm sai trend lớp trên; pie phá ordered time"],
            ["Locate outliers", "Chỉ bảng sort; chỉ line", "Bảng không hiện correlation; line giấu lệch đa biến"],
          ],
          [2400, 3000, 3960]
        ),
        spacer(),
        p(
          "Hạn chế còn lại (ghi nhận, không giấu): kênh size yếu hơn position nên mưa TB các tỉnh khá gần nhau, bubble ít phân hóa; scatter ~20k điểm vẫn overplot (đã tách màu outlier và hạ opacity); heatmap dùng tổng mưa tháng nên tháng 31 ngày hơi nặng hơn tháng 28 — tradeoff của aggregation. Zoom geo đã tắt để tránh mất khung Việt Nam."
        ),

        h2("6.4. Tầng algorithm"),
        p(
          "Dashboard vẽ lại khi lọc; máy không WebGL sẽ chậm scatter. Không suy từ “load nhanh” ra kết luận “vis tốt”. Threat đã xử lý: zoomInGeo làm văng bản đồ — đã gỡ nút zoom geo và khóa khung lat/lon Việt Nam; layout dùng minmax(0,1fr) để card không tràn khi thu cửa sổ."
        ),

        h2("6.5. Kết luận đánh giá"),
        p(
          "Ba idiom khớp ba {action, target} trên slide lý thuyết; linked interaction cho phép chuyển summarize → identify mà không đổi dataset. Pattern khí hậu Việt Nam đọc được trên vis, nên nhóm kết luận thiết kế đạt yêu cầu phân tích–thiết kế–cài đặt. Hướng mở: GeoJSON choropleth đã khớp tên tỉnh; calendar heatmap; user study nhỏ theo insight-based evaluation (North)."
        ),

        h1("7. Phân công thực hiện"),
        table(
          ["Thành viên", "Phân tích và thiết kế", "Cài đặt trên dashboard"],
          [
            [
              "Nguyễn Quốc Trạng\n25C12012",
              "Data abstraction không gian; task discover distribution; encoding map + histogram",
              "Scattergeo, histogram, KPI, bộ lọc liên kết, khung trang",
            ],
            [
              "Trương Thị Ngọc Viên\n25C12038",
              "Task compare trends; chọn line vs heatmap; derive MA 7 ngày",
              "Multi-line, heatmap vùng × tháng, overlay tỉnh đang khóa",
            ],
            [
              "Phạm Quốc Vương\n25C12039",
              "Task locate outliers; quy tắc IQR + z-score; boxplot địa hình",
              "Scattergl, boxplot, bảng 15 outlier, lasso brush",
            ],
          ],
          [2400, 3600, 3360]
        ),
        spacer(),
        p("Cả nhóm cùng thống nhất data abstraction dùng chung và idiom tương tác tích hợp, để ba task không thành ba dashboard rời."),

        h1("8. Link dashboard và cách chạy"),
        p(
          [
            run("Link (dán vào ô nộp / trang bìa): ", { bold: true }),
            new ExternalHyperlink({
              children: [run(DASHBOARD, { color: BLUE, underline: true })],
              link: DASHBOARD,
            }),
          ]
        ),
        p("Cách chạy trên máy chấm / máy nhóm:"),
        numbered("Cài Node.js (nhóm dùng bản v22)."),
        numbered("Mở thư mục dashboard."),
        numbered([run("Chạy "), run("node serve.js", { italics: true }), run(".")]),
        numbered([run("Mở trình duyệt tại "), run(DASHBOARD, { color: BLUE }), run(".")]),
        numbered("Để có URL public trước khi nộp: đẩy thư mục dashboard lên GitHub Pages hoặc Netlify Drop, rồi thay link ở trang bìa và mục này."),
        p("Luồng dữ liệu: df_weather_fixed_utf8.csv → scripts/prepare-data.js → dashboard/data/weather.json → dashboard/index.html."),
        p(
          "Gợi ý kiểm tra tương tác (phục vụ chấm điểm bonus): đổi khoảng ngày và vùng → mọi chart đổi đồng thời; click một tỉnh trên bản đồ → histogram/trend/scatter thu về tỉnh đó; lasso vài điểm đỏ trên scatter → bản đồ chỉ còn các tỉnh tương ứng; bấm “Bỏ chọn tỉnh” / “Đặt lại toàn bộ” để về overview."
        ),

        h1("Tài liệu tham khảo"),
        numbered(
          [
            run("Munzner, T. (2014). "),
            run("Visualization Analysis and Design", { italics: true }),
            run(". CRC Press. Chương 2–5 (what / why / how / nested model)."),
          ],
          "refs"
        ),
        numbered(
          [
            run("Munzner, T. (2009). A Nested Model for Visualization Design and Validation. "),
            run("IEEE TVCG", { italics: true }),
            run(" 15(6), 921–928."),
          ],
          "refs"
        ),
        numbered(
          [
            run("Brehmer, M., & Munzner, T. (2013). A Multi-Level Typology of Abstract Visualization Tasks. "),
            run("IEEE TVCG", { italics: true }),
            run(" 19(12), 2376–2385."),
          ],
          "refs"
        ),
        numbered(
          [
            run("Mackinlay, J. (1986). Automating the Design of Graphical Presentations of Relational Information. "),
            run("ACM TOG", { italics: true }),
            run("."),
          ],
          "refs"
        ),
        numbered(
          [
            run("Shneiderman, B. (1996). The Eyes Have It: A Task by Data Type Taxonomy for Information Visualizations."),
          ],
          "refs"
        ),
        numbered(
          [
            run("Wickham, H., & Stryjewski, L. (2012). 40 years of boxplots."),
          ],
          "refs"
        ),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log("Wrote", OUT, "bytes", buf.length);
});
