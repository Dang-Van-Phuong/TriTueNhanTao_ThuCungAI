// server.js (ES Modules)
// YÊU CẦU: trong package.json có "type":"module" hoặc đổi tên file thành server.mjs

import express from "express";
import nodemailer from "nodemailer";
import dotenv from "dotenv";
import cors from "cors";
import XLSX from "xlsx";

dotenv.config();

const app = express();
app.use(cors({ origin: true }));
app.use(express.json({ limit: "10mb" })); // nhận dữ liệu JSON lớn

/*
Body JSON mong đợi từ front-end:

{
  "student": { "name": "Nguyễn A", "class": "10A1", "id": "", "parentPhone": "", "school": "", "note": "" },
  "summary": { "model": "A (Chat + Face)", "startISO": "2025-01-01T12:00:00.000Z", "durationSec": 180,
               "E": 70, "C": 62, "B": 66, "PSI": 66, "badge": "Ổn định" },
  "sources": {
    "chat":  { "E": 68, "C": 64, "B": 72, "note": "8 tin nhắn | +2/-0 từ cảm xúc" },
    "voice": { "E": 65, "C": 70, "B": 70, "note": "HAPPY (p≈82%)" },
    "face":  { "E": 60, "C": 50, "B": 50, "note": "NEUTRAL (p≈48%)" }
  },
  "segments": [ { "idx":1, "from":"2025-01-01T12:00:00.000Z", "to":"2025-01-01T12:00:30.000Z", "E":66, "C":62, "B":64, "w":0.82 }, ... ],
  "samples":  [ { "t":"2025-01-01T12:00:01.000Z", "E":66, "C":61, "B":63, "w":0.80 }, ... ],
  "toEmail": "tuychinh@truong.edu.vn" // (tuỳ chọn) nếu muốn gửi khác TO_EMAIL trong .env
}
*/

app.post("/api/send-report", async (req, res) => {
  try {
    const {
      student = {},
      summary = {},
      sources = {},
      segments = [],
      samples = [],
      toEmail
    } = req.body || {};

    if (!student?.name || !student?.class) {
      return res.status(400).json({ ok: false, error: "Thiếu tên hoặc lớp học viên." });
    }

    // ===== Sheet 1: PSI (tổng quan)
    const overviewAOA = [
      ["BÁO CÁO TƯ VẤN TÂM LÝ – Súp AI Pet"], [""],
      ["Họ và tên", student.name],
      ["Lớp", student.class],
      ["Mã HS", student.id || ""],
      ["Trường", student.school || ""],
      ["SĐT Phụ huynh", student.parentPhone || ""],
      ["Ghi chú", student.note || ""],
      [""],
      ["Mô hình", summary.model || ""],
      ["Bắt đầu (ISO)", summary.startISO || ""],
      ["Thời lượng (giây)", summary.durationSec ?? ""],
      ["Tổng hợp E", summary.E ?? ""],
      ["Tổng hợp C", summary.C ?? ""],
      ["Tổng hợp B", summary.B ?? ""],
      ["PSI", summary.PSI ?? ""],
      ["Phân loại", summary.badge || ""],
      [""],
      ["Nguồn dữ liệu", "E", "C", "B", "Ghi chú"],
      ["Chat",  sources?.chat?.E  ?? "", sources?.chat?.C  ?? "", sources?.chat?.B  ?? "", sources?.chat?.note  || ""],
      ["Voice", sources?.voice?.E ?? "", sources?.voice?.C ?? "", sources?.voice?.B ?? "", sources?.voice?.note || ""],
      ["Face",  sources?.face?.E  ?? "", sources?.face?.C  ?? "", sources?.face?.B  ?? "", sources?.face?.note  || ""],
    ];
    const wsPSI = XLSX.utils.aoa_to_sheet(overviewAOA);

    // ===== Sheet 2: Phân đoạn 30s
    const segRows = (segments || []).map(seg => ({
      idx: seg.idx,
      from_ISO: seg.from || "",
      to_ISO: seg.to || "",
      E: seg.E ?? "",
      C: seg.C ?? "",
      B: seg.B ?? "",
      w: typeof seg.w === "number" ? Number(seg.w.toFixed(3)) : ""
    }));
    const wsSeg = XLSX.utils.json_to_sheet(segRows, {
      header: ["idx","from_ISO","to_ISO","E","C","B","w"]
    });

    // ===== Sheet 3: Mẫu 1s
    const sampleRows = (samples || []).map((s, i) => ({
      STT: i + 1,
      ISO_Time: s.t || "",
      E: s.E ?? "",
      C: s.C ?? "",
      B: s.B ?? "",
      w: typeof s.w === "number" ? Number(s.w.toFixed(3)) : ""
    }));
    const wsSam = XLSX.utils.json_to_sheet(sampleRows, {
      header: ["STT","ISO_Time","E","C","B","w"]
    });

    // ===== Workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsPSI, "PSI");
    XLSX.utils.book_append_sheet(wb, wsSeg, "Phan_doan_30s");
    XLSX.utils.book_append_sheet(wb, wsSam, "Mau_1s");
    const buf = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });

    // ===== Gửi email
    if (!process.env.GMAIL_USER || !process.env.GMAIL_PASS) {
      return res.status(500).json({ ok: false, error: "Thiếu GMAIL_USER/GMAIL_PASS trong .env" });
    }

    const transporter = nodemailer.createTransport({
      service: "gmail",
      auth: { user: process.env.GMAIL_USER, pass: process.env.GMAIL_PASS }, // App Password
    });

    const recipients = (toEmail && String(toEmail).trim().length > 0)
      ? toEmail
      : (process.env.TO_EMAIL || process.env.GMAIL_USER);

    const subject = `Báo cáo tâm lý – ${student.name} – PSI=${summary?.PSI ?? "?"}`;
    const text =
`Báo cáo mới từ Súp AI Pet

Học viên: ${student.name}
Lớp: ${student.class}
Mã HS: ${student.id || ""}
Trường: ${student.school || ""}
SĐT PH: ${student.parentPhone || ""}
Ghi chú: ${student.note || ""}

Mô hình: ${summary.model || ""}
Bắt đầu: ${summary.startISO || ""}
Thời lượng: ${summary.durationSec ?? ""} giây
Tổng hợp: E=${summary.E ?? ""} | C=${summary.C ?? ""} | B=${summary.B ?? ""}
PSI: ${summary.PSI ?? ""} – ${summary.badge || ""}

File Excel đính kèm gồm 3 sheet: PSI, Phan_doan_30s, Mau_1s.`;

    await transporter.sendMail({
      from: `"Súp AI Pet" <${process.env.GMAIL_USER}>`,
      to: recipients,
      subject,
      text,
      attachments: [{
        filename: `BaoCao_${(student.name || "HocVien").replace(/\s+/g, "_")}.xlsx`,
        content: buf,
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }],
    });

    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: "Gửi email thất bại." });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log(`Mailer server listening on :${port}`));
