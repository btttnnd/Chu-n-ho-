/* =========================================================================
   Add-in Chuẩn thể thức Nghị định 30/2020/NĐ-CP
   =========================================================================
   Ghi chú: đây là công cụ hỗ trợ, tự động hoá phần lớn các yêu cầu về
   font chữ, cỡ chữ, canh lề, cấu trúc thành phần thể thức. Một vài chi tiết
   rất đặc thù (ví dụ đường kẻ ngang có độ dài đúng bằng dòng chữ
   "Độc lập - Tự do - Hạnh phúc") được xấp xỉ bằng gạch chân văn bản; nếu
   đơn vị bạn yêu cầu tuyệt đối chính xác, chỉnh tay thêm sau khi chèn mẫu.
   ========================================================================= */

const MM_TO_PT = 2.834645669;

Office.onReady(() => {
  setupTabs();
  setupCreateTab();
  setupCheckTab();
  setupToolsTab();
});

/* ---------------------------- UI: tabs ---------------------------- */

function setupTabs() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(btn.dataset.tab).classList.add("active");
    });
  });

  document.getElementById("loaiVanBan").addEventListener("change", (e) => {
    const trichYeuField = document.getElementById("fieldTrichYeu");
    // Mọi loại văn bản đều dùng trích yếu, kể cả công văn (ghi dưới dạng "V/v ...")
    trichYeuField.style.display = "block";
  });
}

function setStatus(msg, isError) {
  const bar = document.getElementById("statusBar");
  bar.textContent = msg || "";
  bar.style.color = isError ? "#b3261e" : "";
}

/* ============================================================
   TAB 1: TẠO VĂN BẢN MẪU MỚI
   ============================================================ */

function setupCreateTab() {
  document.getElementById("btnCreate").addEventListener("click", onCreateTemplate);
}

function getCreateFormData() {
  const loaiVanBan = document.getElementById("loaiVanBan").value;
  const coQuanChuQuan = document.getElementById("coQuanChuQuan").value.trim();
  const coQuanBanHanh = document.getElementById("coQuanBanHanh").value.trim();
  const soHieu = document.getElementById("soHieu").value.trim() || "…";
  const kyHieuCoQuan = document.getElementById("kyHieuCoQuan").value.trim() || "…";
  const diaDanh = document.getElementById("diaDanh").value.trim();
  const trichYeu = document.getElementById("trichYeu").value.trim();
  const noiNhanRaw = document.getElementById("noiNhan").value;
  const chucVuKy = document.getElementById("chucVuKy").value.trim();

  const noiNhan = noiNhanRaw
    .split("\n")
    .map((s) => s.trim())
    .filter((s) => s.length > 0);

  return {
    loaiVanBan,
    coQuanChuQuan,
    coQuanBanHanh,
    soHieu,
    kyHieuCoQuan,
    diaDanh,
    trichYeu,
    noiNhan,
    chucVuKy,
  };
}

const TEN_LOAI_VAN_BAN = {
  CONG_VAN: null, // công văn không ghi "tên loại" phía trên trích yếu
  TO_TRINH: "TỜ TRÌNH",
  BAO_CAO: "BÁO CÁO",
  THONG_BAO: "THÔNG BÁO",
  QUYET_DINH: "QUYẾT ĐỊNH",
  KE_HOACH: "KẾ HOẠCH",
  BIEN_BAN: "BIÊN BẢN",
  CONG_DIEN: "CÔNG ĐIỆN",
};

const KY_HIEU_LOAI = {
  CONG_VAN: "CV",
  TO_TRINH: "TTr",
  BAO_CAO: "BC",
  THONG_BAO: "TB",
  QUYET_DINH: "QĐ",
  KE_HOACH: "KH",
  BIEN_BAN: "BB",
  CONG_DIEN: "CĐ",
};

async function onCreateTemplate() {
  const data = getCreateFormData();

  if (!data.coQuanBanHanh || !data.diaDanh || !data.trichYeu || !data.chucVuKy) {
    setStatus("Vui lòng điền đủ các trường có dấu (*).", true);
    return;
  }

  try {
    setStatus("Đang chèn văn bản mẫu…");
    await Word.run(async (context) => {
      const body = context.document.body;

      // 1. Đặt trang & lề chuẩn trước
      applyPageSetup(context, 20, 20, 30, 15);

      // 2. Xoá nội dung hiện có? -> Không xoá, chèn ở đầu tài liệu để an toàn.
      // Chèn từ vị trí đầu tài liệu, theo thứ tự ngược (mỗi lần insert tại Start
      // sẽ đẩy nội dung trước đó xuống dưới), nên ta build toàn bộ khối rồi
      // chèn 1 lần bằng cách chèn vào 1 đoạn trống đầu tiên.

      const firstPara = body.paragraphs.getFirst();
      const anchor = firstPara.insertParagraph("", Word.InsertLocation.before);
      anchor.font.set({ name: "Times New Roman", size: 14 });

      let cursor = anchor;

      // ---- Khối Quốc hiệu / Cơ quan ban hành (bảng 1 dòng x 2 cột, ẩn viền) ----
      const table = cursor.insertTable(1, 2, Word.InsertLocation.after, [["", ""]]);
      table.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
      // Ẩn toàn bộ viền bảng
      ["Top", "Bottom", "Left", "Right", "InsideHorizontal", "InsideVertical"].forEach((edge) => {
        const b = table.getBorder(edge);
        b.type = "None";
      });
      table.getRange().font.set({ name: "Times New Roman" });

      const leftCell = table.getCell(0, 0).body;
      const rightCell = table.getCell(0, 1).body;

      // Cột trái: cơ quan chủ quản (nếu có) + cơ quan ban hành + gạch ngang ngắn + số hiệu
      leftCell.clear();
      if (data.coQuanChuQuan) {
        const p1 = leftCell.insertParagraph(data.coQuanChuQuan.toUpperCase(), Word.InsertLocation.start);
        p1.alignment = Word.Alignment.centered;
        p1.font.set({ bold: false, size: 13, name: "Times New Roman" });
      }
      const p2 = leftCell.insertParagraph(data.coQuanBanHanh.toUpperCase(), Word.InsertLocation.end);
      p2.alignment = Word.Alignment.centered;
      p2.font.set({ bold: true, size: 13, name: "Times New Roman" });

      const soHieuText = `Số: ${data.soHieu}/${data.kyHieuCoQuan}-${KY_HIEU_LOAI[data.loaiVanBan]}`;
      const p3 = leftCell.insertParagraph(soHieuText, Word.InsertLocation.end);
      p3.alignment = Word.Alignment.centered;
      p3.font.set({ bold: false, size: 13, name: "Times New Roman" });

      // Cột phải: Quốc hiệu + Tiêu ngữ (gạch chân xấp xỉ) + địa danh, ngày tháng
      rightCell.clear();
      const q1 = rightCell.insertParagraph(
        "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM",
        Word.InsertLocation.start
      );
      q1.alignment = Word.Alignment.centered;
      q1.font.set({ bold: true, size: 13, name: "Times New Roman" });

      const q2 = rightCell.insertParagraph("Độc lập - Tự do - Hạnh phúc", Word.InsertLocation.end);
      q2.alignment = Word.Alignment.centered;
      q2.font.set({ bold: true, size: 14, name: "Times New Roman", underline: "Single" });

      const today = new Date();
      const ngayThangText = `${data.diaDanh}, ngày ${today.getDate()} tháng ${
        today.getMonth() + 1
      } năm ${today.getFullYear()}`;
      const q3 = rightCell.insertParagraph(ngayThangText, Word.InsertLocation.end);
      q3.alignment = Word.Alignment.centered;
      q3.font.set({ bold: false, italic: true, size: 13, name: "Times New Roman" });

      await context.sync();

      // Word luôn tự thêm một đoạn trống ngay sau bảng; lấy đoạn đó làm điểm chèn tiếp theo
      cursor = table.getRange(Word.RangeLocation.after).paragraphs.getFirst();

      // ---- Tên loại văn bản + trích yếu ----
      const tenLoai = TEN_LOAI_VAN_BAN[data.loaiVanBan];
      if (tenLoai) {
        const pLoai = cursor.insertParagraph(tenLoai, Word.InsertLocation.after);
        pLoai.alignment = Word.Alignment.centered;
        pLoai.font.set({ bold: true, size: 14, name: "Times New Roman" });
        cursor = pLoai;

        const pTrich = cursor.insertParagraph(data.trichYeu, Word.InsertLocation.after);
        pTrich.alignment = Word.Alignment.centered;
        // Trích yếu (văn bản có tên loại): đứng, đậm, cỡ 14, có gạch chân xấp xỉ đường kẻ dưới trích yếu
        pTrich.font.set({ bold: true, italic: false, underline: "Single", size: 14, name: "Times New Roman" });
        cursor = pTrich;
      } else {
        // Công văn: "V/v ..." nằm ngay dưới số ký hiệu, căn giữa, in nghiêng
        const pTrich = cursor.insertParagraph(`V/v ${data.trichYeu}`, Word.InsertLocation.after);
        pTrich.alignment = Word.Alignment.centered;
        // Trích yếu công văn: chữ ĐỨNG (không nghiêng), KHÔNG đậm, cỡ 12 theo đúng NĐ 30
        pTrich.font.set({ bold: false, italic: false, underline: "Single", size: 12, name: "Times New Roman" });
        cursor = pTrich;
      }

      // ---- Dòng trống trước khi vào nội dung ----
      const pBlank1 = cursor.insertParagraph("", Word.InsertLocation.after);
      cursor = pBlank1;

      // ---- Kính gửi (chỉ với công văn) ----
      if (data.loaiVanBan === "CONG_VAN") {
        const pKinhGui = cursor.insertParagraph("Kính gửi: ……………………………………..", Word.InsertLocation.after);
        pKinhGui.font.set({ size: 14, name: "Times New Roman" });
        cursor = pKinhGui;
        const pBlank2 = cursor.insertParagraph("", Word.InsertLocation.after);
        cursor = pBlank2;
      }

      // ---- Placeholder nội dung ----
      const pNoiDung = cursor.insertParagraph(
        "……………………………………………………………………………………………………………………………………………………………………………………………………………………………………………………………………………………………………………..",
        Word.InsertLocation.after
      );
      pNoiDung.font.set({ size: 14, name: "Times New Roman" });
      pNoiDung.alignment = Word.Alignment.justified;
      cursor = pNoiDung;

      const pBlank3 = cursor.insertParagraph("", Word.InsertLocation.after);
      cursor = pBlank3;

      // ---- Khối Nơi nhận (trái) + Chức vụ - ký tên (phải), dùng bảng ẩn viền ----
      const table2 = cursor.insertTable(1, 2, Word.InsertLocation.after, [["", ""]]);
      table2.styleBuiltIn = Word.BuiltInStyleName.tableGrid;
      ["Top", "Bottom", "Left", "Right", "InsideHorizontal", "InsideVertical"].forEach((edge) => {
        table2.getBorder(edge).type = "None";
      });

      const nnCell = table2.getCell(0, 0).body;
      const kyCell = table2.getCell(0, 1).body;

      nnCell.clear();
      const pNN = nnCell.insertParagraph("Nơi nhận:", Word.InsertLocation.start);
      pNN.font.set({ italic: true, bold: true, size: 12, name: "Times New Roman" });
      data.noiNhan.forEach((line) => {
        const p = nnCell.insertParagraph(`- ${line}`, Word.InsertLocation.end);
        // Theo NĐ 30: phần liệt kê Nơi nhận là chữ ĐỨNG (không nghiêng), cỡ 11
        p.font.set({ italic: false, bold: false, size: 11, name: "Times New Roman" });
      });

      kyCell.clear();
      const pChucVu1 = kyCell.insertParagraph(
        data.coQuanChuQuan ? "TM. " + data.coQuanBanHanh.toUpperCase() : "",
        Word.InsertLocation.start
      );
      pChucVu1.alignment = Word.Alignment.centered;
      pChucVu1.font.set({ bold: true, italic: false, size: 14, name: "Times New Roman" });

      const pChucVu2 = kyCell.insertParagraph(data.chucVuKy.toUpperCase(), Word.InsertLocation.end);
      pChucVu2.alignment = Word.Alignment.centered;
      pChucVu2.font.set({ bold: true, italic: false, size: 14, name: "Times New Roman" });

      for (let i = 0; i < 4; i++) {
        kyCell.insertParagraph("", Word.InsertLocation.end);
      }
      // Họ và tên người ký: chữ THƯỜNG (không phải in hoa), đứng, đậm, cỡ 13-14 theo NĐ 30
      const pHoTen = kyCell.insertParagraph("(Họ và tên)", Word.InsertLocation.end);
      pHoTen.alignment = Word.Alignment.centered;
      pHoTen.font.set({ italic: false, bold: true, size: 14, name: "Times New Roman" });

      await context.sync();
    });
    setStatus("Đã chèn văn bản mẫu theo thể thức NĐ 30.");
  } catch (err) {
    console.error(err);
    setStatus("Có lỗi khi chèn văn bản: " + (err.message || err), true);
  }
}

/* ============================================================
   TAB 2: RÀ SOÁT & SỬA VĂN BẢN HIỆN CÓ
   ============================================================ */

function setupCheckTab() {
  document.getElementById("btnFixMargin").addEventListener("click", onFixMargin);
  document.getElementById("btnFixFont").addEventListener("click", onFixFont);
  document.getElementById("btnFixQuocHieu").addEventListener("click", onFixQuocHieu);
  document.getElementById("btnScan").addEventListener("click", onScanAll);
  document.getElementById("btnFixAll").addEventListener("click", onFixAll);
}

/* ---------------------------------------------------------------------
   Bộ phân loại vai trò từng đoạn văn (dùng chung cho sửa toàn bộ & rà soát)
   --------------------------------------------------------------------- */

const NOI_NHAN_ITEM_PREFIX = /^[-–•]\s*/;
const DATE_LINE_REGEX = /^(.*?),\s*ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})\s*\.?$/i;
const SO_HIEU_REGEX = /^Số\s*:\s*\S+/i;
const CAN_CU_REGEX = /^Căn\s+cứ\s+/i;
const TEN_LOAI_VB_LIST = [
  "CÔNG VĂN", "TỜ TRÌNH", "BÁO CÁO", "THÔNG BÁO", "QUYẾT ĐỊNH",
  "KẾ HOẠCH", "BIÊN BẢN", "CÔNG ĐIỆN", "CHỈ THỊ", "NGHỊ QUYẾT", "CÔNG BỐ",
];

// Dò trước 2 dòng "Cơ quan chủ quản" / "Cơ quan ban hành" (nếu có) — đây là các dòng in hoa
// ngắn nằm NGAY TRƯỚC dòng Quốc hiệu, nên cần biết trước vị trí Quốc hiệu mới phân biệt được
// với dòng "chức vụ ký" (cũng in hoa ngắn nhưng nằm ở cuối văn bản).
function detectHeaderOrgRoles(paragraphTexts) {
  const map = new Map();
  const quocHieuIndex = paragraphTexts.findIndex((t) => {
    const u = (t || "").toUpperCase();
    return u.includes("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM") || u.includes("CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM");
  });
  if (quocHieuIndex <= 0) return map;

  const hasLowerVN = (t) =>
    /[a-zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]/.test(t);

  const candidates = [];
  for (let i = 0; i < quocHieuIndex; i++) {
    const t = normalizeForMatch(paragraphTexts[i]);
    if (!t) continue;
    if (!hasLowerVN(t) && t.length >= 3 && t.length <= 80) {
      candidates.push(i);
    }
  }
  const lastTwo = candidates.slice(-2);
  if (lastTwo.length === 1) {
    map.set(lastTwo[0], "CO_QUAN_BAN_HANH");
  } else if (lastTwo.length === 2) {
    map.set(lastTwo[0], "CO_QUAN_CHU_QUAN");
    map.set(lastTwo[1], "CO_QUAN_BAN_HANH");
  }
  return map;
}

// state = { insideNoiNhan: bool, afterChucVu: bool } — truyền cùng 1 object qua các lần gọi
// liên tiếp theo đúng thứ tự đoạn văn trong tài liệu để theo dõi khối Nơi nhận và dòng Họ tên
// ngay sau dòng chức vụ ký.
function normalizeForMatch(s) {
  return (s || "")
    .replace(/[\u00A0\u200B\u202F\u2007\u2060]/g, " ") // các loại khoảng trắng đặc biệt -> khoảng trắng thường
    .replace(/\s+/g, " ")
    .trim();
}

function classifyParagraph(rawText, state) {
  const text = normalizeForMatch(rawText);
  if (!text) {
    state.insideNoiNhan = false;
    return "EMPTY";
  }
  const upper = text.toUpperCase();

  if (upper.includes("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM") || upper.includes("CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM")) {
    state.insideNoiNhan = false;
    return "QUOC_HIEU";
  }
  if (upper.includes("ĐỘC LẬP") && upper.includes("TỰ DO") && upper.includes("HẠNH PHÚC")) {
    state.insideNoiNhan = false;
    return "TIEU_NGU";
  }
  if (DATE_LINE_REGEX.test(text)) {
    state.insideNoiNhan = false;
    return "NGAY_THANG";
  }
  if (SO_HIEU_REGEX.test(text)) {
    state.insideNoiNhan = false;
    return "SO_HIEU";
  }
  if (CAN_CU_REGEX.test(text)) {
    state.insideNoiNhan = false;
    return "CAN_CU";
  }
  if (/^V\/v\s+/i.test(text)) {
    state.insideNoiNhan = false;
    return "TRICH_YEU_CONGVAN";
  }
  if (text.toLowerCase().startsWith("nơi nhận")) {
    state.insideNoiNhan = true;
    return "NOI_NHAN_LABEL";
  }
  if (state.insideNoiNhan && NOI_NHAN_ITEM_PREFIX.test(text)) {
    return "NOI_NHAN_ITEM";
  }
  // Ra khỏi khối Nơi nhận nếu dòng hiện tại không còn bắt đầu bằng dấu gạch đầu dòng
  state.insideNoiNhan = false;

  if (TEN_LOAI_VB_LIST.includes(upper)) {
    state.afterTenLoai = true;
    return "TEN_LOAI_VB";
  }

  // Trích yếu nội dung: dòng NGAY SAU dòng tên loại văn bản (VD sau "BÁO CÁO", "QUYẾT ĐỊNH"...)
  if (state.afterTenLoai) {
    state.afterTenLoai = false; // chỉ xét đúng 1 dòng kế tiếp
    return "TRICH_YEU_AUTO";
  }

  // Chức vụ ký: toàn chữ hoa (không có ký tự thường), ngắn, không phải Quốc hiệu/tên loại
  const hasLowerVietnamese = /[a-zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]/.test(
    text
  );
  if (!hasLowerVietnamese && text.length <= 60 && text.length >= 3) {
    state.afterChucVu = 10; // dò rộng hơn qua nhiều dòng trống (khoảng trống chờ đóng dấu/ký tay)
    return "CHUC_VU_KY";
  }

  // Họ và tên người ký: dò trong vài dòng nội dung kế tiếp sau dòng chức vụ (bỏ qua các dòng
  // trống hoặc dòng không giống tên người, dành cho khoảng trống đóng dấu/ký tay) — chữ thường,
  // 2-5 từ đều viết hoa chữ cái đầu, hoặc placeholder "(Họ và tên)".
  if (state.afterChucVu > 0) {
    if (isLikelySignerName(text)) {
      state.afterChucVu = 0;
      return "SIGNER_NAME";
    }
    state.afterChucVu -= 1;
  }

  return "BODY";
}

const SIGNER_TITLE_PREFIXES = [
  "Đại tướng", "Thượng tướng", "Trung tướng", "Thiếu tướng",
  "Đại tá", "Thượng tá", "Trung tá", "Thiếu tá",
  "Đại úy", "Thượng úy", "Trung úy", "Thiếu úy", "Chuẩn úy",
  "Bác sĩ", "Tiến sĩ", "Thạc sĩ", "Kỹ sư", "Luật sư", "Dược sĩ", "Cử nhân",
];

function isLikelySignerName(text) {
  let t = normalizeForMatch(text).replace(/[.,;:]+$/, "").trim();
  if (/^\(?\s*họ\s+(và\s+)?tên\s*\)?$/i.test(t)) return true;

  // Bỏ qua cấp bậc/danh xưng đứng trước tên (VD "Trung tá Phạm Đức Vinh") khi kiểm tra
  for (const prefix of SIGNER_TITLE_PREFIXES) {
    if (t.startsWith(prefix + " ")) {
      t = t.slice(prefix.length).trim();
      break;
    }
  }

  const words = t.split(/\s+/);
  if (words.length < 2 || words.length > 5) return false;
  const wordPattern =
    /^[A-ZÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ][a-zàáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]*$/;
  return words.every((w) => wordPattern.test(w));
}

function formatDateLineText(text) {
  const m = text.match(DATE_LINE_REGEX);
  if (!m) return null;
  const diaDanh = m[1];
  let day = parseInt(m[2], 10);
  let month = parseInt(m[3], 10);
  const year = m[4];
  const dayStr = day < 10 ? "0" + day : "" + day;
  // Theo NĐ 30: chỉ tháng 1 và 2 mới bắt buộc thêm số 0 phía trước
  const monthStr = month === 1 || month === 2 ? "0" + month : "" + month;
  return `${diaDanh}, ngày ${dayStr} tháng ${monthStr} năm ${year}`;
}

function applyPageSetup(context, topMm, bottomMm, leftMm, rightMm) {
  const section = context.document.sections.getFirst();
  section.pageSetup.topMargin = topMm * MM_TO_PT;
  section.pageSetup.bottomMargin = bottomMm * MM_TO_PT;
  section.pageSetup.leftMargin = leftMm * MM_TO_PT;
  section.pageSetup.rightMargin = rightMm * MM_TO_PT;
}

const CM_TO_PT = 28.3464567;

function xmlEscape(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Tạo 1 đoạn văn OOXML canh giữa, có đường kẻ ngang THẬT (border) bên dưới — dùng thay cho
// gạch chân bằng định dạng chữ, đúng tinh thần "đường kẻ ngang, nét liền" của NĐ 30.
// sizeHalfPt = cỡ chữ * 2 (đơn vị OOXML là nửa điểm).
function buildBorderedCenteredParagraphOoxml(text, sizeHalfPt, bold) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
    "<pkg:xmlData>" +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    "<w:body><w:p>" +
    "<w:pPr>" +
    '<w:jc w:val="center"/>' +
    '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>' +
    "</w:pPr>" +
    "<w:r><w:rPr>" +
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>' +
    (bold ? "<w:b/>" : "") +
    '<w:sz w:val="' +
    sizeHalfPt +
    '"/><w:szCs w:val="' +
    sizeHalfPt +
    '"/>' +
    "</w:rPr><w:t xml:space=\"preserve\">" +
    xmlEscape(text) +
    "</w:t></w:r>" +
    "</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>"
  );
}

async function onFixAll() {
  const fontSize = parseInt(document.getElementById("fontSizeAll").value, 10) || 14;
  const top = parseFloat(document.getElementById("mgTopAll").value) || 20;
  const bottom = parseFloat(document.getElementById("mgBottomAll").value) || 20;
  const left = parseFloat(document.getElementById("mgLeftAll").value) || 30;
  const right = parseFloat(document.getElementById("mgRightAll").value) || 15;

  const counts = {
    QUOC_HIEU: 0,
    TIEU_NGU: 0,
    NGAY_THANG: 0,
    NOI_NHAN_LABEL: 0,
    NOI_NHAN_ITEM: 0,
    CHUC_VU_KY: 0,
    SIGNER_NAME: 0,
    TEN_LOAI_VB: 0,
    TRICH_YEU_AUTO: 0,
    TRICH_YEU_CONGVAN: 0,
    CO_QUAN_CHU_QUAN: 0,
    CO_QUAN_BAN_HANH: 0,
    SO_HIEU: 0,
    CAN_CU: 0,
    BODY: 0,
  };
  let tableCount = 0;
  void tableCount;

  try {
    setStatus("Đang rà soát & sửa toàn bộ văn bản, vui lòng đợi…");
    await Word.run(async (context) => {
      // 1. Khổ giấy A4 + lề trang
      const section = context.document.sections.getFirst();
      section.pageSetup.pageWidth = 210 * MM_TO_PT;
      section.pageSetup.pageHeight = 297 * MM_TO_PT;
      applyPageSetup(context, top, bottom, left, right);

      // 2. Duyệt toàn bộ đoạn văn, phân loại & sửa theo từng vai trò
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();

      const paragraphTexts = paragraphs.items.map((p) => (p.text || "").trim());
      const orgRoleMap = detectHeaderOrgRoles(paragraphTexts);

      const state = { insideNoiNhan: false };

      // NĐ 30: Quốc hiệu/Tiêu ngữ/địa danh-ngày tháng phải dùng cỡ chữ ĐỒNG BỘ theo 1 trong 2 cặp:
      // (Quốc hiệu 13 - Tiêu ngữ 14 - ngày tháng 14) hoặc (Quốc hiệu 12 - Tiêu ngữ 13 - ngày tháng 13)
      const isSmallScale = fontSize === 13;
      const quocHieuSize = isSmallScale ? 12 : 13;
      const tieuNguSize = isSmallScale ? 13 : 14;
      const ngayThangSize = tieuNguSize;

      paragraphs.items.forEach((p, idx) => {
        const text = paragraphTexts[idx];
        let role = classifyParagraph(text, state);
        if (orgRoleMap.has(idx)) role = orgRoleMap.get(idx);
        if (counts[role] !== undefined) counts[role]++;

        // Các vai trò cần canh giữa TRÊN TOÀN CHIỀU RỘNG trang — phải xoá sạch mọi thụt lề
        // (left/right/first-line indent) kế thừa từ định dạng gốc, nếu không "canh giữa" sẽ
        // bị tính lệch theo phần còn lại sau khi trừ thụt lề, nhìn như chưa canh giữa.
        const CENTERED_ROLES = new Set([
          "CO_QUAN_CHU_QUAN", "CO_QUAN_BAN_HANH", "SO_HIEU", "CAN_CU",
          "QUOC_HIEU", "TIEU_NGU", "NGAY_THANG",
          "CHUC_VU_KY", "SIGNER_NAME", "TEN_LOAI_VB", "TRICH_YEU_AUTO", "TRICH_YEU_CONGVAN",
        ]);
        if (CENTERED_ROLES.has(role)) {
          p.leftIndent = 0;
          p.rightIndent = 0;
          p.firstLineIndent = 0;
          p.alignment = Word.Alignment.centered;
        }

        switch (role) {
          case "CO_QUAN_CHU_QUAN":
            // Cơ quan chủ quản: in hoa, đứng, KHÔNG đậm theo NĐ 30
            p.font.set({ name: "Times New Roman", size: quocHieuSize, bold: false, italic: false, underline: "None" });
            break;

          case "CO_QUAN_BAN_HANH":
            // Cơ quan ban hành: in hoa, đứng, ĐẬM, có gạch chân xấp xỉ đường kẻ dưới (yêu cầu
            // thực tế là 1/3-1/2 độ dài dòng chữ, nhưng gạch chân theo chữ chỉ làm được full dòng)
            p.font.set({ name: "Times New Roman", size: quocHieuSize, bold: true, italic: false, underline: "Single" });
            break;

          case "SO_HIEU":
            // Số, ký hiệu văn bản: chữ đứng, không đậm, cỡ 13, canh giữa
            p.font.set({ name: "Times New Roman", size: 13, bold: false, italic: false });
            break;

          case "CAN_CU":
            // Căn cứ ban hành: chữ nghiêng, không đậm, canh giữa, cùng cỡ chữ nội dung
            p.font.set({ name: "Times New Roman", size: fontSize, bold: false, italic: true });
            break;

          case "QUOC_HIEU":
            p.font.set({ name: "Times New Roman", size: quocHieuSize, bold: true, italic: false, underline: "None" });
            break;

          case "TIEU_NGU":
            // Đã thử đường kẻ thật (OOXML) 2 lần nhưng không chèn được trên môi trường Word của
            // người dùng — quay lại cách an toàn, luôn hoạt động: gạch chân bằng định dạng chữ.
            p.font.set({ name: "Times New Roman", size: tieuNguSize, bold: true, italic: false, underline: "Single" });
            break;

          case "NGAY_THANG": {
            const fixed = formatDateLineText(text);
            if (fixed && fixed !== text) {
              p.clear();
              p.insertText(fixed, Word.InsertLocation.start);
            }
            p.font.set({ name: "Times New Roman", size: ngayThangSize, bold: false, italic: true });
            break;
          }

          case "NOI_NHAN_LABEL":
            p.font.set({ name: "Times New Roman", size: 12, bold: true, italic: true });
            p.alignment = Word.Alignment.left;
            break;

          case "NOI_NHAN_ITEM":
            // Theo đúng NĐ 30: phần liệt kê là chữ ĐỨNG (không nghiêng), cỡ chữ 11, không đậm
            p.font.set({ name: "Times New Roman", size: 11, bold: false, italic: false });
            p.alignment = Word.Alignment.left;
            break;

          case "CHUC_VU_KY":
            // Quyền hạn/chức vụ: in hoa, đứng, đậm, cỡ 13-14 theo NĐ 30 — ép cứng về 14 để đảm bảo
            // luôn nằm trong khung cho phép (không giữ nguyên cỡ sai của tài liệu gốc)
            p.font.set({ name: "Times New Roman", size: 14, bold: true, italic: false });
            break;

          case "SIGNER_NAME":
            // Họ và tên người ký: chữ thường, đứng, đậm, cỡ 13-14 theo NĐ 30 — ép cứng về 14
            p.font.set({ name: "Times New Roman", size: 14, bold: true, italic: false });
            break;

          case "TEN_LOAI_VB":
            p.font.set({ name: "Times New Roman", size: 14, bold: true, italic: false });
            break;

          case "TRICH_YEU_AUTO":
            // Trích yếu (văn bản có tên loại): chữ thường, đứng, đậm, cỡ theo cỡ nội dung, canh giữa
            p.font.set({ name: "Times New Roman", size: fontSize, bold: true, italic: false });
            break;

          case "TRICH_YEU_CONGVAN":
            // Trích yếu công văn (dòng "V/v..."): chữ đứng, không đậm, cỡ 12, canh giữa
            p.font.set({ name: "Times New Roman", size: 12, bold: false, italic: false });
            break;

          case "BODY":
            p.font.set({ name: "Times New Roman", size: fontSize });
            p.alignment = Word.Alignment.justified;
            p.firstLineIndent = CM_TO_PT * 1; // thụt đầu dòng 1cm
            p.lineSpacing = 18; // tương đương giãn dòng 1.5
            p.spaceAfter = 6;
            p.spaceBefore = 0;
            break;

          default:
            break; // EMPTY: bỏ qua
        }
      });

      await context.sync();
    });

    const parts = [];
    if (counts.CO_QUAN_CHU_QUAN) parts.push("cơ quan chủ quản");
    if (counts.CO_QUAN_BAN_HANH) parts.push("cơ quan ban hành");
    if (counts.SO_HIEU) parts.push("số hiệu văn bản");
    if (counts.CAN_CU) parts.push(`${counts.CAN_CU} căn cứ ban hành`);
    if (counts.QUOC_HIEU) parts.push("Quốc hiệu");
    if (counts.TIEU_NGU) parts.push("Tiêu ngữ");
    if (counts.NGAY_THANG) parts.push("ngày tháng năm");
    if (counts.NOI_NHAN_LABEL || counts.NOI_NHAN_ITEM) parts.push("Nơi nhận");
    if (counts.CHUC_VU_KY) parts.push("chức vụ ký");
    if (counts.SIGNER_NAME) parts.push("họ tên người ký");
    if (counts.TEN_LOAI_VB) parts.push("tên loại văn bản");
    if (counts.TRICH_YEU_AUTO || counts.TRICH_YEU_CONGVAN) parts.push("trích yếu");
    if (counts.BODY) parts.push(`${counts.BODY} đoạn nội dung`);
    if (tableCount) parts.push(`${tableCount} bảng`);

    setStatus(
      `Đã sửa xong theo NĐ 30${parts.length ? ": " + parts.join(", ") : ""}. Khổ giấy A4, lề trên ${top}mm/dưới ${bottom}mm/trái ${left}mm/phải ${right}mm.`
    );
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi sửa toàn bộ: " + (err.message || err), true);
  }
}

async function onFixMargin() {
  const top = parseFloat(document.getElementById("mgTop").value) || 20;
  const bottom = parseFloat(document.getElementById("mgBottom").value) || 20;
  const left = parseFloat(document.getElementById("mgLeft").value) || 30;
  const right = parseFloat(document.getElementById("mgRight").value) || 15;

  try {
    setStatus("Đang áp dụng khổ giấy & lề trang…");
    await Word.run(async (context) => {
      const section = context.document.sections.getFirst();
      section.pageSetup.pageWidth = 210 * MM_TO_PT; // A4
      section.pageSetup.pageHeight = 297 * MM_TO_PT;
      applyPageSetup(context, top, bottom, left, right);
      await context.sync();
    });
    setStatus(`Đã đặt lề: trên ${top}mm, dưới ${bottom}mm, trái ${left}mm, phải ${right}mm (A4).`);
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi áp dụng lề trang: " + (err.message || err), true);
  }
}

async function onFixFont() {
  const size = parseInt(document.getElementById("fontSize").value, 10) || 14;
  try {
    setStatus("Đang áp dụng font chữ…");
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();

      const state = { insideNoiNhan: false };
      paragraphs.items.forEach((p) => {
        const text = (p.text || "").trim();
        const role = classifyParagraph(text, state);
        // Chỉ đổi cỡ chữ cho đoạn nội dung thường (BODY); các khối Nơi nhận, chữ ký,
        // Quốc hiệu... giữ nguyên cỡ đặc thù của chúng, chỉ đồng bộ tên font.
        if (role === "BODY" || role === "EMPTY") {
          p.font.set({ name: "Times New Roman", size: size });
        } else {
          p.font.name = "Times New Roman";
        }
      });
      await context.sync();
    });
    setStatus(`Đã áp dụng Times New Roman cho toàn văn bản (nội dung cỡ ${size}, giữ nguyên cỡ riêng của Nơi nhận/chữ ký/Quốc hiệu).`);
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi áp dụng font: " + (err.message || err), true);
  }
}

async function onFixQuocHieu() {
  try {
    setStatus("Đang rà soát Quốc hiệu, Tiêu ngữ…");
    let found1 = false;
    let found2 = false;
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();

      for (const p of paragraphs.items) {
        const text = (p.text || "").trim();
        const upper = text.toUpperCase();

        if (upper.includes("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")) {
          p.alignment = Word.Alignment.centered;
          p.font.set({ name: "Times New Roman", size: 13, bold: true });
          found1 = true;
        }

        if (
          upper.includes("ĐỘC LẬP") &&
          upper.includes("TỰ DO") &&
          upper.includes("HẠNH PHÚC")
        ) {
          p.alignment = Word.Alignment.centered;
          p.font.set({ name: "Times New Roman", size: 14, bold: true, underline: "Single" });
          found2 = true;
        }
      }
      await context.sync();
    });

    if (found1 || found2) {
      setStatus(
        `Đã sửa: ${found1 ? "Quốc hiệu" : ""}${found1 && found2 ? " & " : ""}${
          found2 ? "Tiêu ngữ" : ""
        }.`
      );
    } else {
      setStatus(
        "Không tìm thấy dòng Quốc hiệu/Tiêu ngữ trong văn bản. Dùng tab 'Tạo văn bản mới' nếu muốn chèn mới.",
        true
      );
    }
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi rà soát Quốc hiệu/Tiêu ngữ: " + (err.message || err), true);
  }
}

async function onScanAll() {
  const reportEl = document.getElementById("scanReport");
  reportEl.innerHTML = "";
  const items = [];

  try {
    setStatus("Đang rà soát…");
    await Word.run(async (context) => {
      const section = context.document.sections.getFirst();
      section.pageSetup.load("topMargin,bottomMargin,leftMargin,rightMargin,pageWidth,pageHeight");

      const body = context.document.body;
      body.font.load("name,size");

      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text,items/font/name,items/font/size");

      await context.sync();

      // 1. Lề trang
      const top = section.pageSetup.topMargin / MM_TO_PT;
      const bottom = section.pageSetup.bottomMargin / MM_TO_PT;
      const left = section.pageSetup.leftMargin / MM_TO_PT;
      const right = section.pageSetup.rightMargin / MM_TO_PT;

      const marginOk =
        top >= 19.5 && top <= 25.5 &&
        bottom >= 19.5 && bottom <= 25.5 &&
        left >= 29.5 && left <= 35.5 &&
        right >= 14.5 && right <= 20.5;

      items.push({
        ok: marginOk,
        text: marginOk
          ? `Lề trang hợp lệ (trên ${top.toFixed(1)}mm, dưới ${bottom.toFixed(1)}mm, trái ${left.toFixed(
              1
            )}mm, phải ${right.toFixed(1)}mm).`
          : `Lề trang chưa đúng khung NĐ 30 (hiện: trên ${top.toFixed(1)}mm, dưới ${bottom.toFixed(
              1
            )}mm, trái ${left.toFixed(1)}mm, phải ${right.toFixed(
              1
            )}mm). Yêu cầu: trên/dưới 20-25mm, trái 30-35mm, phải 15-20mm.`,
      });

      // 2. Khổ giấy A4
      const pw = section.pageSetup.pageWidth / MM_TO_PT;
      const ph = section.pageSetup.pageHeight / MM_TO_PT;
      const isA4 = Math.abs(pw - 210) < 3 && Math.abs(ph - 297) < 3;
      items.push({
        ok: isA4,
        text: isA4
          ? "Khổ giấy A4 hợp lệ."
          : `Khổ giấy hiện tại ${pw.toFixed(0)}x${ph.toFixed(0)}mm, không phải A4 (210x297mm).`,
      });

      // 3. Font & cỡ chữ trong các đoạn văn (bỏ qua đoạn trống)
      let wrongFontCount = 0;
      let checkedCount = 0;
      paragraphs.items.forEach((p) => {
        const text = (p.text || "").trim();
        if (!text) return;
        checkedCount++;
        const fname = p.font.name;
        const fsize = p.font.size;
        const nameOk = fname && fname.toLowerCase().includes("times new roman");
        const sizeOk = fsize === 13 || fsize === 14 || fsize === 12 || fsize === 12.5;
        if (!nameOk || !sizeOk) wrongFontCount++;
      });
      items.push({
        ok: wrongFontCount === 0,
        text:
          wrongFontCount === 0
            ? `Font chữ hợp lệ trên ${checkedCount} đoạn có nội dung.`
            : `Có ${wrongFontCount}/${checkedCount} đoạn dùng font/cỡ chữ khác Times New Roman 12-14. Bấm "Áp dụng" ở mục Font chữ để sửa nhanh.`,
      });

      // 4. Quốc hiệu, tiêu ngữ có tồn tại không
      const hasQuocHieu = paragraphs.items.some((p) =>
        (p.text || "").toUpperCase().includes("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
      );
      const hasTieuNgu = paragraphs.items.some((p) => {
        const u = (p.text || "").toUpperCase();
        return u.includes("ĐỘC LẬP") && u.includes("TỰ DO") && u.includes("HẠNH PHÚC");
      });
      items.push({
        ok: hasQuocHieu && hasTieuNgu,
        text:
          hasQuocHieu && hasTieuNgu
            ? "Đã có đủ Quốc hiệu và Tiêu ngữ."
            : `Thiếu: ${!hasQuocHieu ? "Quốc hiệu" : ""}${
                !hasQuocHieu && !hasTieuNgu ? " và " : ""
              }${!hasTieuNgu ? "Tiêu ngữ" : ""}. Dùng tab "Tạo văn bản mới" để chèn khối chuẩn, hoặc gõ tay rồi bấm "Rà soát & sửa" ở mục Quốc hiệu.`,
      });

      // 5. Nơi nhận
      const hasNoiNhan = paragraphs.items.some((p) => (p.text || "").trim().toLowerCase().startsWith("nơi nhận"));
      items.push({
        ok: hasNoiNhan,
        text: hasNoiNhan ? "Đã có mục Nơi nhận." : "Chưa thấy mục 'Nơi nhận' trong văn bản.",
      });
    });

    // Render report
    items.forEach((it) => {
      const div = document.createElement("div");
      div.className = "report-item " + (it.ok ? "ok" : "warn");
      div.textContent = (it.ok ? "✓ " : "⚠ ") + it.text;
      reportEl.appendChild(div);
    });

    setStatus("Rà soát xong.");
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi rà soát: " + (err.message || err), true);
  }
}

/* ============================================================
   TAB 3: TIỆN ÍCH
   ============================================================ */

function setupToolsTab() {
  document.getElementById("btnInsertLandscape").addEventListener("click", () => onInsertOrientedPage("Landscape"));
  document.getElementById("btnInsertPortrait").addEventListener("click", () => onInsertOrientedPage("Portrait"));
  document.getElementById("btnRemoveBlankPages").addEventListener("click", onRemoveBlankPages);
  document.getElementById("btnInsertPageNumber").addEventListener("click", onInsertPageNumber);
  document.getElementById("btnQuickSize13").addEventListener("click", () => onQuickSize(13));
  document.getElementById("btnQuickSize14").addEventListener("click", () => onQuickSize(14));
  document.getElementById("btnQuickSize15").addEventListener("click", () => onQuickSize(15));
  document.getElementById("btnKeepWithNext").addEventListener("click", onKeepWithNext);
  document.getElementById("btnRepeatTableHeader").addEventListener("click", onRepeatTableHeader);
  document.getElementById("btnCenterTables").addEventListener("click", onCenterTables);
  document.getElementById("btnCenterImages").addEventListener("click", onCenterImages);
  document.getElementById("btnCleanExcelPaste").addEventListener("click", onCleanExcelPaste);
  document.getElementById("btnCellAlignTop").addEventListener("click", () => onCellAlign("Top"));
  document.getElementById("btnCellAlignCenter").addEventListener("click", () => onCellAlign("Center"));
  document.getElementById("btnInsertQr").addEventListener("click", onInsertQr);
  document.getElementById("btnRemoveDiacritics").addEventListener("click", onRemoveDiacritics);
  document.getElementById("btnDecimalComma").addEventListener("click", onDecimalComma);
}

async function onInsertOrientedPage(orientation) {
  try {
    setStatus(`Đang chèn trang ${orientation === "Landscape" ? "ngang" : "dọc"}…`);
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);
      await context.sync();

      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      const lastSection = sections.items[sections.items.length - 1];
      lastSection.pageSetup.load("pageWidth,pageHeight,orientation");
      await context.sync();

      const w = lastSection.pageSetup.pageWidth;
      const h = lastSection.pageSetup.pageHeight;
      lastSection.pageSetup.orientation = orientation;
      // Đảm bảo khổ giấy xoay đúng chiều (một số phiên bản Word không tự hoán đổi W/H)
      if (orientation === "Landscape" && w < h) {
        lastSection.pageSetup.pageWidth = h;
        lastSection.pageSetup.pageHeight = w;
      } else if (orientation === "Portrait" && w > h) {
        lastSection.pageSetup.pageWidth = h;
        lastSection.pageSetup.pageHeight = w;
      }
      await context.sync();
    });
    setStatus(`Đã chèn trang ${orientation === "Landscape" ? "ngang" : "dọc"} ở cuối văn bản.`);
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi chèn trang: " + (err.message || err), true);
  }
}

async function onRemoveBlankPages() {
  try {
    setStatus("Đang xoá đoạn trống thừa ở cuối văn bản…");
    let removed = 0;
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();

      const items = paragraphs.items;
      // Xoá các đoạn trống liên tiếp ở cuối văn bản, luôn giữ lại ít nhất 1 đoạn
      let i = items.length - 1;
      while (i > 0 && (items[i].text || "").trim() === "") {
        items[i].delete();
        removed++;
        i--;
      }
      await context.sync();
    });
    if (removed > 0) {
      setStatus(`Đã xoá ${removed} đoạn trống thừa ở cuối văn bản (nguyên nhân phổ biến nhất gây trang trắng thừa).`);
    } else {
      setStatus("Không tìm thấy đoạn trống thừa ở cuối văn bản.");
    }
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi xoá trang thừa: " + (err.message || err), true);
  }
}

async function onInsertPageNumber() {
  try {
    setStatus("Đang chèn số trang…");
    await Word.run(async (context) => {
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();

      sections.items.forEach((sec) => {
        const footer = sec.getFooter(Word.HeaderFooterType.primary);
        footer.clear();
        const ooxmlField =
          '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"></pkg:package>';
        // Chèn field PAGE bằng OOXML tối giản (fldSimple) để số trang tự cập nhật
        footer.insertOoxml(
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
            '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
            "<pkg:xmlData>" +
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
            "<w:body><w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>" +
            '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
            "</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>",
          Word.InsertLocation.end
        );
      });
      await context.sync();
    });
    setStatus("Đã chèn số trang (canh giữa) vào chân trang.");
  } catch (err) {
    console.error(err);
    setStatus(
      "Lỗi khi chèn số trang (một số phiên bản Word chặn chèn OOXML dạng field): " + (err.message || err),
      true
    );
  }
}

async function onQuickSize(size) {
  try {
    setStatus(`Đang chuyển cỡ chữ nội dung sang ${size}…`);
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();

      const state = { insideNoiNhan: false };
      paragraphs.items.forEach((p) => {
        const text = (p.text || "").trim();
        const role = classifyParagraph(text, state);
        if (role === "BODY" || role === "EMPTY") {
          p.font.size = size;
        }
      });
      await context.sync();
    });
    setStatus(`Đã chuyển cỡ chữ nội dung sang ${size} (giữ nguyên Quốc hiệu/Nơi nhận/chữ ký).`);
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi đổi cỡ chữ: " + (err.message || err), true);
  }
}

async function onKeepWithNext() {
  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      const paragraphs = sel.paragraphs;
      paragraphs.load("items");
      await context.sync();
      paragraphs.items.forEach((p) => {
        p.keepWithNext = true;
      });
      await context.sync();
    });
    setStatus("Đã bật 'Keep with next' cho đoạn đang chọn.");
  } catch (err) {
    console.error(err);
    setStatus("Lỗi: " + (err.message || err), true);
  }
}

async function onRepeatTableHeader() {
  try {
    setStatus("Đang thiết lập lặp tiêu đề bảng…");
    let count = 0;
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();
      tables.items.forEach((t) => {
        t.headerRowCount = 1;
        count++;
      });
      await context.sync();
    });
    setStatus(`Đã đặt lặp dòng tiêu đề cho ${count} bảng.`);
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi lặp tiêu đề bảng: " + (err.message || err), true);
  }
}

async function onCenterTables() {
  try {
    setStatus("Đang căn giữa các bảng…");
    let count = 0;
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();
      tables.items.forEach((t) => {
        // Lưu ý: Word JS API không có thuộc tính "table.alignment" để căn giữa cả bảng trên
        // trang — chỉ AutoFit theo cửa sổ (Word sẽ tự canh giữa khi bảng vừa khít nội dung)
        t.autoFitWindow();
        count++;
      });
      await context.sync();
    });
    setStatus(`Đã AutoFit ${count} bảng theo chiều rộng trang.`);
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi căn giữa bảng: " + (err.message || err), true);
  }
}

async function onCenterImages() {
  try {
    setStatus("Đang căn giữa & chỉnh vừa ảnh…");
    let count = 0;
    await Word.run(async (context) => {
      const pics = context.document.body.inlinePictures;
      pics.load("items");
      await context.sync();

      const section = context.document.sections.getFirst();
      section.pageSetup.load("pageWidth,leftMargin,rightMargin");
      await context.sync();
      const contentWidth = section.pageSetup.pageWidth - section.pageSetup.leftMargin - section.pageSetup.rightMargin;

      pics.items.forEach((pic) => {
        pic.lockAspectRatio = true;
        if (pic.width > contentWidth) {
          pic.width = contentWidth;
        }
        const parentPara = pic.paragraph;
        parentPara.alignment = Word.Alignment.centered;
        count++;
      });
      await context.sync();
    });
    setStatus(`Đã căn giữa & chỉnh vừa chiều ngang cho ${count} ảnh.`);
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi căn giữa ảnh: " + (err.message || err), true);
  }
}

async function onCleanExcelPaste() {
  try {
    setStatus("Đang dọn ký tự thừa trong các ô bảng…");
    let count = 0;
    await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();

      // Duyệt từng ô: xoá khoảng trắng/ký tự điều khiển thừa ở đầu-cuối, gộp dòng trống thừa
      for (const t of tables.items) {
        t.load("rowCount,values");
      }
      await context.sync();

      for (const t of tables.items) {
        const rowCount = t.rowCount;
        for (let r = 0; r < rowCount; r++) {
          const cellCount = t.values[r].length;
          for (let c = 0; c < cellCount; c++) {
            const cell = t.getCell(r, c);
            const raw = t.values[r][c] || "";
            const cleaned = raw
              .replace(/[\u200B\u00A0\t]+/g, " ")
              .replace(/\s+\n/g, "\n")
              .replace(/\n{2,}/g, "\n")
              .trim();
            if (cleaned !== raw) {
              cell.body.clear();
              cell.body.insertText(cleaned, Word.InsertLocation.start);
              count++;
            }
          }
        }
      }
      await context.sync();
    });
    setStatus(`Đã dọn ${count} ô bảng có ký tự thừa (khoảng trắng lạ, dòng trống thừa).`);
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi dọn bảng: " + (err.message || err), true);
  }
}

async function onCellAlign(vAlign) {
  try {
    setStatus("Đang căn chỉnh ô…");
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      const table = sel.parentTableOrNullObject;
      table.load("isNullObject");
      await context.sync();

      if (table.isNullObject) {
        setStatus("Vui lòng đặt con trỏ vào trong 1 bảng trước khi bấm nút này.", true);
        return;
      }

      table.load("rowCount,values");
      await context.sync();
      const rowCount = table.rowCount;
      for (let r = 0; r < rowCount; r++) {
        const cellCount = table.values[r].length;
        for (let c = 0; c < cellCount; c++) {
          const cell = table.getCell(r, c);
          cell.verticalAlignment = vAlign === "Top" ? Word.VerticalAlignment.top : Word.VerticalAlignment.center;
        }
      }
      await context.sync();
    });
    setStatus(`Đã căn ${vAlign === "Top" ? "đỉnh" : "giữa"} ô cho bảng đang chọn.`);
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi căn ô: " + (err.message || err), true);
  }
}

async function onInsertQr() {
  const text = document.getElementById("qrText").value.trim();
  if (!text) {
    setStatus("Vui lòng nhập nội dung cho mã QR.", true);
    return;
  }
  try {
    setStatus("Đang tạo mã QR…");
    const url = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(text)}`;
    const resp = await fetch(url);
    if (!resp.ok) throw new Error("Không tải được ảnh mã QR (kiểm tra kết nối mạng).");
    const blob = await resp.blob();
    const base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result.split(",")[1]);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });

    await Word.run(async (context) => {
      context.document.body.insertInlinePictureFromBase64(base64, Word.InsertLocation.end);
      await context.sync();
    });
    setStatus("Đã chèn mã QR vào cuối văn bản.");
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi chèn mã QR: " + (err.message || err), true);
  }
}

const VN_DIACRITICS_MAP = [
  [/[àáạảãâầấậẩẫăằắặẳẵ]/g, "a"],
  [/[ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴ]/g, "A"],
  [/[èéẹẻẽêềếệểễ]/g, "e"],
  [/[ÈÉẸẺẼÊỀẾỆỂỄ]/g, "E"],
  [/[ìíịỉĩ]/g, "i"],
  [/[ÌÍỊỈĨ]/g, "I"],
  [/[òóọỏõôồốộổỗơờớợởỡ]/g, "o"],
  [/[ÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠ]/g, "O"],
  [/[ùúụủũưừứựửữ]/g, "u"],
  [/[ÙÚỤỦŨƯỪỨỰỬỮ]/g, "U"],
  [/[ỳýỵỷỹ]/g, "y"],
  [/[ỲÝỴỶỸ]/g, "Y"],
  [/đ/g, "d"],
  [/Đ/g, "D"],
];

function stripVietnameseDiacritics(text) {
  let result = text;
  VN_DIACRITICS_MAP.forEach(([re, rep]) => {
    result = result.replace(re, rep);
  });
  return result;
}

async function onRemoveDiacritics() {
  try {
    setStatus("Đang xoá dấu tiếng Việt…");
    let count = 0;
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();

      paragraphs.items.forEach((p) => {
        const text = p.text || "";
        const stripped = stripVietnameseDiacritics(text);
        if (stripped !== text) {
          p.clear();
          p.insertText(stripped, Word.InsertLocation.start);
          count++;
        }
      });
      await context.sync();
    });
    setStatus(`Đã xoá dấu cho ${count} đoạn văn. Lưu ý: thao tác này làm mất định dạng chi tiết (in đậm/nghiêng từng phần) của các đoạn bị đổi — nên dùng trên bản sao.`);
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi xoá dấu: " + (err.message || err), true);
  }
}

function convertNumberToVnStyle(text) {
  // Chỉ đổi các cụm số kiểu Anh-Mỹ rõ ràng: có dấu phẩy phân nghìn (>=1 nhóm 3 số)
  // và tuỳ chọn phần thập phân sau dấu chấm. Vd: 1,234,567.89 -> 1.234.567,89
  return text.replace(/\b\d{1,3}(,\d{3})+(\.\d+)?\b/g, (match) => {
    const withPlaceholder = match.replace(/,/g, "§");
    const withComma = withPlaceholder.replace(/\./g, ",");
    return withComma.replace(/§/g, ".");
  });
}

async function onDecimalComma() {
  try {
    setStatus("Đang đổi định dạng số sang kiểu Việt Nam…");
    let count = 0;
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text");
      await context.sync();

      paragraphs.items.forEach((p) => {
        const text = p.text || "";
        const converted = convertNumberToVnStyle(text);
        if (converted !== text) {
          p.clear();
          p.insertText(converted, Word.InsertLocation.start);
          count++;
        }
      });
      await context.sync();
    });
    if (count > 0) {
      setStatus(`Đã đổi định dạng số ở ${count} đoạn văn (chỉ đổi số có dấu phẩy phân nghìn kiểu Anh-Mỹ).`);
    } else {
      setStatus("Không tìm thấy số nào theo định dạng Anh-Mỹ (1,234.56) để đổi.");
    }
  } catch (err) {
    console.error(err);
    setStatus("Lỗi khi đổi định dạng số: " + (err.message || err), true);
  }
}
