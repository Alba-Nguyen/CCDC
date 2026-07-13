import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const load = async (p) => SpreadsheetFile.importXlsx(await FileBlob.load(p));
const template = await load("Check_CCDC 1.xlsx");
const bkctWb = await load("Bảng kê chứng từ.xlsx");
const bkpbWb = await load("Bảng tổng hợp chi phí chờ phân bổ.xlsx");

// Always read the complete used range. Source reports vary in row count by period.
const bkctSrc = bkctWb.worksheets.getItemAt(0).getUsedRange().values;
const bkpbSrc = bkpbWb.worksheets.getItemAt(0).getUsedRange().values;

// Normalize the two new reports to the exact column structure used by Check_CCDC 1.
// BKCT: remove blank merged-description column D and blank column W.
const bkct = bkctSrc.map(r => r.filter((_, i) => i !== 3 && i !== 22));
// BKPB: remove blank merged-name column C and blank separator column M.
const bkpb = bkpbSrc.map(r => r.filter((_, i) => i !== 2 && i !== 12));

const shCt = template.worksheets.getItem("02.BKCT");
const shPb = template.worksheets.getItem("03.BKPB");
const shCheck = template.worksheets.getItem("CheckChiTiet");

// The new source reports already contain their own title text in cells; remove the
// old template text boxes to prevent duplicated/overlapping headings.
shCt.deleteAllDrawings();
shPb.deleteAllDrawings();

const oldCtRows = shCt.getUsedRange().values.length;
const oldPbRows = shPb.getUsedRange().values.length;
shCt.getRange(`A1:AB${Math.max(oldCtRows, bkct.length)}`).clear({applyTo:"contents"});
shCt.getRange(`A1:X${bkct.length}`).values = bkct;
shPb.getRange(`A1:M${Math.max(oldPbRows, bkpb.length)}`).clear({applyTo:"contents"});
shPb.getRange(`A1:M${bkpb.length}`).values = bkpb;

// Identify actual transaction/detail rows.
let ctLast = 8;
for (let i = 8; i < bkct.length; i++) if (typeof bkct[i][5] === "number" && bkct[i][2]) ctLast = i + 1;
let pbLast = 5;
// Cột A (Mã) có thể để trống ở một số dòng hợp lệ. Xác định dòng cuối
// theo diễn giải (B) hoặc số tiền (F) để luôn lấy hết dữ liệu nguồn.
for (let i = 5; i < bkpb.length; i++) {
  const hasDescription = String(bkpb[i][1] ?? "").trim() !== "";
  const hasAmount = typeof bkpb[i][5] === "number";
  if (hasDescription || hasAmount) pbLast = i + 1;
}

// Account-aware lookup key prevents identical names in 2421 and 2422 from being
// assigned to the wrong CCDC code.
shPb.getRange("N5").values = [["Key tra Mã"]];
shPb.getRange(`N6:N${pbLast}`).values = bkpb.slice(5,pbLast).map(r => [[`${r[1] ?? ""}|${r[9] ?? ""}`][0]]);

shCt.getRange("Y8:AB8").values = [["Xuly_1","Xuly_2","Code CCDC","Tên CCDC"]];
shCt.getRange("Y9").formulas = [["=_xlfn.TEXTAFTER($C9,\":\")"]];
shCt.getRange(`Y9:Y${ctLast}`).fillDown();
shCt.getRange("Z9").formulas = [["=IF(Y9=\"\",\"\",RIGHT(Y9,8))"]];
shCt.getRange(`Z9:Z${ctLast}`).fillDown();
shCt.getRange("AA9").formulas = [["=_xlfn.XLOOKUP($AB9,'03.BKPB'!$B:$B,'03.BKPB'!$A:$A,\"\",0,1)"]];
shCt.getRange(`AA9:AA${ctLast}`).fillDown();
const accountNameToCode = new Map();
for (let i = 5; i < pbLast; i++) {
  const key = `${bkpb[i][1] ?? ""}|${bkpb[i][9] ?? ""}`;
  if (!accountNameToCode.has(key)) accountNameToCode.set(key, bkpb[i][0] ?? "");
}
const mappedCodes = [];
const mappedNames = [];
for (let i = 8; i < ctLast; i++) {
  const desc = String(bkct[i][2] ?? "");
  const after = desc.includes(":") ? desc.split(":").slice(1).join(":").trim() : "";
  const name = after.replace(/\s+T\d{1,2}\.\d{4}\s*$/i, "");
  mappedCodes.push([accountNameToCode.get(`${name}|${bkct[i][4] ?? ""}`) ?? ""]);
  mappedNames.push([name]);
}
shCt.getRange("AC8").values = [["Code theo TK"]];
shCt.getRange(`AC9:AC${ctLast}`).values = mappedCodes;
shCt.getRange("AD8").values = [["Tên đối chiếu"]];
shCt.getRange(`AD9:AD${ctLast}`).values = mappedNames;
// Remove the leading space and month suffix without TRIM, so embedded line breaks
// remain identical to BKPB names and exact-match lookup continues to work.
shCt.getRange("AB9").formulas = [["=TRIM(_xlfn.TEXTBEFORE(Y9,Z9))"]];
shCt.getRange(`AB9:AB${ctLast}`).fillDown();
for (let i = 8; i < ctLast; i++) {
  if (!String(bkct[i][2] ?? "").includes(":")) shCt.getRange(`Y${i+1}:AB${i+1}`).clear({applyTo:"contents"});
}

// Build separate comparison sheets by account so 2421 and 2422 cannot offset each other.
const nameToCode = new Map();
for (let i = 5; i < pbLast; i++) if (bkpb[i][1]) nameToCode.set(String(bkpb[i][1]).trim(), bkpb[i][0] ?? "");
function pairsFor(account) {
  const pairs = new Map();
  for (let i = 5; i < pbLast; i++) {
    if (String(bkpb[i][9] ?? "") === account && bkpb[i][1]) pairs.set(`${bkpb[i][0] ?? ""}\u0000${bkpb[i][1]}`, [bkpb[i][0] ?? "", bkpb[i][1]]);
  }
  for (let i = 8; i < ctLast; i++) {
    if (String(bkct[i][4] ?? "") !== account || !bkct[i][2]) continue;
    const after = String(bkct[i][2]).split(":").slice(1).join(":").trim();
    const name = after.replace(/\s+T\d{1,2}\.\d{4}\s*$/i, "").trim();
    if (!name) continue;
    const code = nameToCode.get(name) ?? "";
    pairs.set(`${code}\u0000${name}`, [code, name]);
  }
  return [...pairs.values()].sort((a,b) => String(a[0]).localeCompare(String(b[0]),"vi") || String(a[1]).localeCompare(String(b[1]),"vi"));
}

const sh21 = template.worksheets.add("CheckChiTiet_2421");
sh21.getRange("A1:G200").copyFrom(shCheck.getRange("A1:G200"), "all");

function populateCheck(sheet, account) {
  const list = pairsFor(account);
  const last = 3 + list.length;
  sheet.getRange("A4:G400").clear({applyTo:"contents"});
  sheet.getRange(`A4:B${last}`).values = list;
  sheet.getRange("D1").values = [[Number(account)]];
  sheet.getRange("C2").formulas = [[`=SUBTOTAL(9,C4:C${last})`]];
  sheet.getRange("D2").formulas = [[`=SUBTOTAL(9,D4:D${last})`]];
  sheet.getRange("E2").formulas = [["=C2-D2"]];
  sheet.getRange("C4").formulas = [[`=IF($A4=\"\",SUMIFS('02.BKCT'!$F$9:$F$${ctLast},'02.BKCT'!$AD$9:$AD$${ctLast},$B4,'02.BKCT'!$E$9:$E$${ctLast},$D$1),SUMIFS('02.BKCT'!$F$9:$F$${ctLast},'02.BKCT'!$AC$9:$AC$${ctLast},$A4,'02.BKCT'!$AD$9:$AD$${ctLast},$B4,'02.BKCT'!$E$9:$E$${ctLast},$D$1))`]];
  sheet.getRange(`C4:C${last}`).fillDown();
  sheet.getRange("D4").formulas = [[`=SUMIFS('03.BKPB'!$F$6:$F$${pbLast},'03.BKPB'!$A$6:$A$${pbLast},$A4,'03.BKPB'!$B$6:$B$${pbLast},$B4,'03.BKPB'!$J$6:$J$${pbLast},$D$1)`]];
  sheet.getRange(`D4:D${last}`).fillDown();
  sheet.getRange("E4").formulas = [["=C4-D4"]];
  sheet.getRange(`E4:E${last}`).fillDown();
  sheet.getRange("F4").formulas = [[`=SUMIFS($E$4:$E$${last},$B$4:$B$${last},B4)`]];
  sheet.getRange(`F4:F${last}`).fillDown();
  sheet.getRange("G4").formulas = [["=IF(F4=0,0,\"Check\")"]];
  sheet.getRange(`G4:G${last}`).fillDown();
  sheet.getRange("H3").values = [["Ghi chú"]];
  sheet.getRange("H4").formulas = [["=IF(A4=\"\",\"Không tìm thấy Mã\",\"\")"]];
  sheet.getRange(`H4:H${last}`).fillDown();
  sheet.getRange("H3").format = {fill:"#FCE4D6",font:{bold:true},horizontalAlignment:"center",verticalAlignment:"center",borders:{preset:"all",style:"thin",color:"#BFBFBF"}};
  sheet.getRange(`H4:H${last}`).format.borders = {preset:"all",style:"thin",color:"#D9D9D9"};
  sheet.getRange("H:H").format.columnWidth = 24;
  sheet.getRange(`C4:F${last}`).format.numberFormat = "#,##0;[Red]-#,##0;0";
  sheet.freezePanes.freezeRows(3);
  return {last, items:list.length};
}
const check22 = populateCheck(shCheck,"2422");
const check21 = populateCheck(sh21,"2421");
shCheck.name = "CheckChiTiet_2422";

const shInf = template.worksheets.add("Inf");
shInf.showGridLines = false;
shInf.getRange("A1:F1").merge();
shInf.getRange("A1").values = [["HƯỚNG DẪN LẤY FILE NGUỒN ĐỐI CHIẾU CCDC"]];
shInf.getRange("A1:F1").format = {fill:"#1F4E78",font:{bold:true,color:"#FFFFFF",size:16},horizontalAlignment:"center",verticalAlignment:"center"};
shInf.getRange("A3:F3").values = [["STT","File nguồn","Đường dẫn trên Bravo 8","Báo cáo cần chọn","Điều kiện/Lưu ý","Tên sheet đích"]];
shInf.getRange("A3:F3").format = {fill:"#D9EAF7",font:{bold:true},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#9EADBA"}};
shInf.getRange("A4:F5").values = [
  [1,"Bảng kê chứng từ","BÁO CÁO → Sổ kế toán theo hình thức nhật ký chung","Bảng kê chứng từ","Lấy dữ liệu tài khoản 2421 và 2422 trong kỳ cần đối chiếu","02.BKCT"],
  [2,"Bảng tổng hợp chi phí chờ phân bổ","BÁO CÁO → Báo cáo tài sản cố định, công cụ dụng cụ","Bảng tổng hợp chi phí chờ phân bổ","BẮT BUỘC tích chọn: “Tích hợp thêm dữ liệu CCDC”","03.BKPB"],
];
shInf.getRange("A4:F5").format = {verticalAlignment:"top",wrapText:true,borders:{preset:"all",style:"thin",color:"#D9D9D9"}};
shInf.getRange("E5").format = {fill:"#FFF2CC",font:{bold:true,color:"#C00000"},wrapText:true,borders:{preset:"all",style:"thin",color:"#D6B656"}};
shInf.getRange("A7:F7").merge();
shInf.getRange("A7").values = [["Sau khi xuất báo cáo, thay dữ liệu vào đúng sheet nguồn; không xóa hoặc sửa các cột công thức của file mẫu."]];
shInf.getRange("A7:F7").format = {fill:"#E2F0D9",font:{italic:true,color:"#375623"},wrapText:true,borders:{preset:"outside",style:"thin",color:"#A9D18E"}};
shInf.getRange("A:A").format.columnWidth = 8;
shInf.getRange("B:B").format.columnWidth = 32;
shInf.getRange("C:C").format.columnWidth = 55;
shInf.getRange("D:D").format.columnWidth = 42;
shInf.getRange("E:E").format.columnWidth = 48;
shInf.getRange("F:F").format.columnWidth = 18;
shInf.getRange("1:1").format.rowHeight = 30;
shInf.getRange("3:3").format.rowHeight = 32;
shInf.getRange("4:5").format.rowHeight = 64;
shInf.getRange("7:7").format.rowHeight = 32;
// Match the visible structure of the template on the added 2421 sheet.
sh21.showGridLines = false;
sh21.getRange("A1:G1").format.font = {bold:true,size:16};
sh21.getRange("A3:B3").format = {fill:"#D9EAF7",font:{bold:true},horizontalAlignment:"center",verticalAlignment:"center",borders:{preset:"all",style:"thin",color:"#BFBFBF"}};
sh21.getRange("C3:E3").format = {fill:"#FFF2CC",font:{bold:true},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#BFBFBF"}};
sh21.getRange("F3:G3").format = {fill:"#E2F0D9",font:{bold:true},horizontalAlignment:"center",verticalAlignment:"center",borders:{preset:"all",style:"thin",color:"#BFBFBF"}};
sh21.getRange("C2:E2").format.font = {bold:true,color:"#FF0000"};
sh21.getRange(`A4:G${check21.last}`).format.borders = {preset:"all",style:"thin",color:"#D9D9D9"};
sh21.getRange("A:A").format.columnWidth = 20;
sh21.getRange("B:B").format.columnWidth = 72;
sh21.getRange("C:D").format.columnWidth = 15;
sh21.getRange("E:F").format.columnWidth = 18;
sh21.getRange("G:G").format.columnWidth = 12;
sh21.getRange("3:3").format.rowHeight = 32;

const out = await SpreadsheetFile.exportXlsx(template);
await out.save("TOOL_Check_CCDC/_temp_stage.xlsx");
console.log(JSON.stringify({ctLast,pbLast,check22,check21,output:"TOOL_Check_CCDC/_temp_stage.xlsx"}));
