import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

/**
 * SBMAB203 文档修正生成工具。
 *
 * 功能：
 * 基于原始 SQL 仕様書与业务处理详细工作簿，按预定义修正清单定点改写单元格，并把改动文字标红后输出新文件。
 *
 * 作用：
 * 在不破坏原始 xlsx 样式、页签结构和布局节奏的前提下，把本轮人工核对出的文档差异沉淀为可重复执行的本地修正工具。
 *
 * 适用场景：
 * - 根据 `SBMAMapper.xml` 修正对应 SQL 仕様書
 * - 根据 `SBMAB203.java` 复核业务处理详细书是否需要调整
 * - 在离线环境下重复生成同格式的修正文档和改动报告
 */
public class GenerateSbma203DocCorrections {

  /** OPTION 根目录。 */
  private static final Path OPTION_ROOT = Paths.get("C:\\opt\\workspace\\SELPLAT\\OPTION");

  /** 原始文档目录。 */
  private static final Path SOURCE_DOC_DIR = OPTION_ROOT.resolve("doc");

  /** 修正文档输出目录。 */
  private static final Path OUTPUT_DOC_DIR = OPTION_ROOT.resolve("generated_doc");

  /** 改动报告路径。 */
  private static final Path REPORT_PATH = OPTION_ROOT.resolve("SBMAB203_doc_correction_report.md");

  /** 红色字体缓存，避免重复创建相同样式。 */
  private final Map<String, CellStyle> redStyleCache = new HashMap<>();

  /** 工作簿计划结果。 */
  private final List<WorkbookResult> workbookResults = new ArrayList<>();

  /**
   * 主入口。
   *
   * @param args 可选参数，当前未使用。
   * @throws Exception 生成过程中任一步骤失败时抛出。
   */
  public static void main(String[] args) throws Exception {
    // 业务上直接执行时默认处理 OPTION/doc 下的既有文档，避免再让调用方传重复路径。
    GenerateSbma203DocCorrections generator = new GenerateSbma203DocCorrections();
    generator.generate();
  }

  /**
   * 统一执行全部文档修正计划。
   *
   * @throws Exception 读取、改写或输出失败时抛出。
   */
  private void generate() throws Exception {
    // 业务上先确保输出目录存在，避免最后写文件时才暴露路径问题。
    Files.createDirectories(OUTPUT_DOC_DIR);

    // SQL 仕様書的修正计划。
    applyWorkbookPlan(buildAgentCompletePlan());
    applyWorkbookPlan(buildIssuerMergePlan());
    applyWorkbookPlan(buildIssuerCompletePlan());
    applyWorkbookPlan(buildNormalMasterCheckOneAndTwoPlan());
    applyWorkbookPlan(buildNormalMasterCheckFourPlan());

    // 业务处理详细书本轮核对未发现需要改动的单元格，但仍把检查结果写入报告。
    workbookResults.add(new WorkbookResult(
        "業務機能（バッチ）処理詳細_SBMAB203_整合性チェック.xlsx",
        "",
        new ArrayList<>(),
        "已核对页签结构、主流程方法族和关键后处理代码，本轮未发现需要修正的文档差异。"));

    // 业务上最终交付不仅要有新文档，还要留下可复核的改动说明。
    writeReport();
  }

  /**
   * 应用单个工作簿计划并输出修正文件。
   *
   * @param plan 单个工作簿的修正计划。
   * @throws Exception 打开、改写或保存失败时抛出。
   */
  private void applyWorkbookPlan(WorkbookPlan plan) throws Exception {
    // 业务上按原文件名精确命中要修正的工作簿，避免误改同目录下其他样本。
    Path sourcePath = SOURCE_DOC_DIR.resolve(plan.sourceFileName());
    Path outputPath = OUTPUT_DOC_DIR.resolve(plan.outputFileName());
    List<ChangeRecord> records = new ArrayList<>();

    try (InputStream inputStream = Files.newInputStream(sourcePath);
        Workbook workbook = WorkbookFactory.create(inputStream)) {
      // 业务上每个工作簿可能有多步插行和改单元格，统一按计划顺序执行。
      for (WorkbookOperation operation : plan.operations()) {
        operation.apply(workbook, records);
      }
      // 输出到新文件，确保原始文档不被覆盖。
      try (OutputStream outputStream = Files.newOutputStream(outputPath)) {
        workbook.write(outputStream);
      }
    }

    workbookResults.add(new WorkbookResult(plan.sourceFileName(), outputPath.toString(), records, ""));
  }

  /**
   * 生成 SB 代理人设置済み检查规格书的修正计划。
   *
   * @return 对应工作簿计划。
   */
  private WorkbookPlan buildAgentCompletePlan() {
    List<WorkbookOperation> operations = new ArrayList<>();
    operations.add(insertStyledRow("SB代理人設定済みチェック2-3取得 ", 92, 91,
        Map.of("C", "AND TO_CHAR(t1.USE_END_YMD_DT,'YYYYMMDD') = #{bizOpeDate}"),
        "补充 SQL 明细中遗漏的代理人利用结束日过滤条件。"));
    operations.add(insertStyledRow("SB代理人設定済みチェック2-3取得 ", 78, 77,
        Map.of("C", "AND t2.APPL_BEGIN_YMD_DT = #{nxtBizD}"),
        "补充 SQL 明细中遗漏的翌营业日适用开始日条件。"));
    operations.add(setCellText("SB代理人設定済みチェック2-3取得 ", "C75", "LEFT OUTER JOIN",
        "修正 SQL 明细中的 JOIN 类型，使其与 mapper 保持一致。"));
    operations.add(setCellText("SB代理人設定済みチェック2-3取得 ", "F5", "SB代理人設定済みチェック2-3取得",
        "修正 SQL 名单元格的 #VALUE! 错误，恢复与 SQLID 对应的正式文档名称。"));
    operations.add(insertStyledRow("SB代理人設定済みチェック2-3取得 ", 36, 35,
        Map.of("C", "3", "E", "翌営業日", "M", "nxtBizD", "Z", "String"),
        "补充 mapper 实际使用但文档缺失的翌营业日参数。"));
    return new WorkbookPlan(
        "SQL仕様書_SB代理人設定済みチェック2-3取得.xlsx",
        "SQL仕様書_SB代理人設定済みチェック2-3取得_修正.xlsx",
        operations);
  }

  /**
   * 生成 SB 发行者合并检查规格书的修正计划。
   *
   * @return 对应工作簿计划。
   */
  private WorkbookPlan buildIssuerMergePlan() {
    List<WorkbookOperation> operations = new ArrayList<>();
    operations.add(setCellText("SB発行者合併チェック1取得", "F5", "SB発行者合併チェック1取得",
        "修正 SQL 名单元格的 #VALUE! 错误，恢复与 SQLID 对应的正式文档名称。"));
    operations.add(setCellText("SB発行者合併チェック1取得", "AI28", "toCharでStringに変換したNo.3の「利用開始年月日」",
        "修正取得项备注，明确 useBeginYmdDt 来源于 t3。"));
    operations.add(setCellText("SB発行者合併チェック1取得", "AI29", "toCharでStringに変換したNo.3の「利用終了年月日」",
        "修正取得项备注，明确 useEndYmdDt 来源于 t3。"));
    operations.add(setCellText("SB発行者合併チェック1取得", "AU47", "t1.新会社コード＝t3.会社コード",
        "修正表关联条件，明确法人表按新公司代码关联。"));
    operations.add(setCellText("SB発行者合併チェック1取得", "C74", "TO_CHAR(t3.USE_BEGIN_YMD_DT,'YYYYMMDD')AS useBeginYmdDt,",
        "修正 SQL 明细，明确利用开始日来自法人表 t3。"));
    operations.add(setCellText("SB発行者合併チェック1取得", "C75", "TO_CHAR(t3.USE_END_YMD_DT,'YYYYMMDD')AS useEndYmdDt,",
        "修正 SQL 明细，明确利用结束日来自法人表 t3。"));
    operations.add(setCellText("SB発行者合併チェック1取得", "C95", "ON t1.NEWCO_CD = t3.CO_CD",
        "修正 SQL 明细中的法人关联条件，使其与 mapper 保持一致。"));
    return new WorkbookPlan(
        "SQL仕様書_SB発行者合併チェック1取得.xlsx",
        "SQL仕様書_SB発行者合併チェック1取得_修正.xlsx",
        operations);
  }

  /**
   * 生成 SB 发行者设置済み检查规格书的修正计划。
   *
   * @return 对应工作簿计划。
   */
  private WorkbookPlan buildIssuerCompletePlan() {
    List<WorkbookOperation> operations = new ArrayList<>();
    operations.add(insertStyledRow("SB発行者設定済みチェック1取得 ", 77, 76,
        Map.of("C", "AND TO_CHAR(t1.USE_END_YMD_DT,'YYYYMMDD') = #{bizOpeDate}"),
        "补充 SQL 明细中遗漏的利用结束日过滤条件。"));
    operations.add(insertStyledRow("SB発行者設定済みチェック1取得 ", 72, 71,
        Map.of("C", "AND t2.APPL_BEGIN_YMD_DT = #{nxtBizD}"),
        "补充 SQL 明细中遗漏的翌营业日适用开始日条件。"));
    operations.add(insertStyledRow("SB発行者設定済みチェック1取得 ", 45, 44,
        Map.of("E", "No.1の「利用終了年月日」がパラメータの「業務運用日」と一致かつ、"),
        "补充抽出条件摘要中遗漏的利用结束日条件。"));
    operations.add(setCellText("SB発行者設定済みチェック1取得 ", "AU40",
        "t1.会社コード＝t2.会社コード、t2.適用開始年月日＝パラメータの「翌営業日」",
        "补充表关联条件中的翌营业日适用开始日限制。"));
    operations.add(insertStyledRow("SB発行者設定済みチェック1取得 ", 32, 31,
        Map.of("C", "2", "E", "翌営業日", "M", "nxtBizD", "Z", "String"),
        "补充 mapper 实际使用但文档缺失的翌营业日参数。"));
    return new WorkbookPlan(
        "SQL仕様書_SB発行者設定済みチェック1取得.xlsx",
        "SQL仕様書_SB発行者設定済みチェック1取得_修正.xlsx",
        operations);
  }

  /**
   * 生成一般债券主表检查 1-2 的修正计划。
   *
   * @return 对应工作簿计划。
   */
  private WorkbookPlan buildNormalMasterCheckOneAndTwoPlan() {
    List<WorkbookOperation> operations = new ArrayList<>();
    operations.add(setCellText("一般債銘柄マスタチェック1-2取得", "E17", "SB資金決済会社/SB資金決済会社(翌営業日)",
        "修正利用表说明，补足 masterFlag=0 时使用的基础表。"));
    operations.add(setCellText("一般債銘柄マスタチェック1-2取得", "V17", "TB_FSBBANK/TB_FSBBANKNXT",
        "修正利用表物理名，补足基础表与翌营业日表的双路径。"));
    return new WorkbookPlan(
        "SQL仕様書_一般債銘柄マスタチェック1-2取得.xlsx",
        "SQL仕様書_一般債銘柄マスタチェック1-2取得_修正.xlsx",
        operations);
  }

  /**
   * 生成一般债券主表检查 4 的修正计划。
   *
   * @return 对应工作簿计划。
   */
  private WorkbookPlan buildNormalMasterCheckFourPlan() {
    List<WorkbookOperation> operations = new ArrayList<>();
    operations.add(insertStyledRow("一般債銘柄マスタチェック4取得", 74, 73,
        Map.of("C", "AND t1.SB_ISSUER_CD IS NOT NULL"),
        "补充 SQL 明细中遗漏的 SB 发行者代码非空条件。"));
    operations.add(insertStyledRow("一般債銘柄マスタチェック4取得", 44, 43,
        Map.of("E", "No.1の「SB発行者コード」が空(null)ではないかつ、"),
        "补充抽出条件摘要中遗漏的非空判断。"));
    operations.add(setCellText("一般債銘柄マスタチェック4取得", "F5", "一般債銘柄マスタチェック4取得",
        "修正 SQL 名单元格的 #VALUE! 错误，恢复与 SQLID 对应的正式文档名称。"));
    operations.add(setCellText("一般債銘柄マスタチェック4取得", "C17", "2",
        "修正利用表编号跳号问题，恢复连续编号。"));
    operations.add(setCellText("一般債銘柄マスタチェック4取得", "AI25", "No.1の「SB発行者コード」",
        "修正取得项备注，把错误的发行业务字段说明改为 SB 发行者代码。"));
    return new WorkbookPlan(
        "SQL仕様書_一般債銘柄マスタチェック4取得.xlsx",
        "SQL仕様書_一般債銘柄マスタチェック4取得_修正.xlsx",
        operations);
  }

  /**
   * 创建简单改单元格操作。
   *
   * @param sheetName 页签名。
   * @param cellRef 单元格坐标。
   * @param newValue 新值。
   * @param reason 变更原因。
   * @return 对应操作。
   */
  private WorkbookOperation setCellText(String sheetName, String cellRef, String newValue, String reason) {
    return (workbook, records) -> {
      // 业务上直接按坐标改单元格，能最大限度保持未改区域的原样式与布局。
      Sheet sheet = workbook.getSheet(sheetName);
      Cell cell = getOrCreateCell(sheet, cellRef);
      String oldValue = readCellText(cell);
      // 把整格新文案改为红色，满足“改动文字标红”的交付要求。
      cell.setCellValue(newValue);
      applyRedStyle(workbook, cell);
      records.add(new ChangeRecord(sheetName, cellRef, oldValue, newValue, reason));
    };
  }

  /**
   * 创建插入样式行并写值的操作。
   *
   * @param sheetName 页签名。
   * @param rowNumberOneBased 插入行号，1 基准。
   * @param templateRowOneBased 复制样式的模板行号，1 基准。
   * @param values 要写入的列和值。
   * @param reason 变更原因。
   * @return 对应操作。
   */
  private WorkbookOperation insertStyledRow(
      String sheetName,
      int rowNumberOneBased,
      int templateRowOneBased,
      Map<String, String> values,
      String reason) {
    return (workbook, records) -> {
      // 业务上先整体下移目标行及其后续内容，再把新条件行插回原位置，避免压坏既有段落。
      Sheet sheet = workbook.getSheet(sheetName);
      int insertRowIndex = rowNumberOneBased - 1;
      sheet.shiftRows(insertRowIndex, sheet.getLastRowNum(), 1, true, false);

      // 新行继承邻近模板行的高度和单元格样式，保证版式节奏不被破坏。
      Row templateRow = sheet.getRow(templateRowOneBased - 1);
      Row insertedRow = sheet.createRow(insertRowIndex);
      copyRowStructure(templateRow, insertedRow);

      // 业务上只写本次新增条件需要的列，其他列保持模板行样式空白。
      for (Map.Entry<String, String> entry : values.entrySet()) {
        String cellRef = entry.getKey() + rowNumberOneBased;
        Cell cell = getOrCreateCell(sheet, cellRef);
        String newValue = entry.getValue();
        cell.setCellValue(newValue);
        applyRedStyle(workbook, cell);
        records.add(new ChangeRecord(sheetName, cellRef, "", newValue, reason));
      }
    };
  }

  /**
   * 复制模板行的结构与样式。
   *
   * @param templateRow 模板行。
   * @param targetRow 新建行。
   */
  private void copyRowStructure(Row templateRow, Row targetRow) {
    // 业务上新插入的行必须尽量贴住周边排版，所以行高与每格样式都直接继承模板行。
    if (templateRow == null) {
      return;
    }
    targetRow.setHeight(templateRow.getHeight());
    short firstCell = templateRow.getFirstCellNum();
    short lastCell = templateRow.getLastCellNum();
    if (firstCell < 0 || lastCell < 0) {
      return;
    }
    for (int cellIndex = firstCell; cellIndex < lastCell; cellIndex++) {
      Cell templateCell = templateRow.getCell(cellIndex);
      if (templateCell == null) {
        continue;
      }
      Cell targetCell = targetRow.createCell(cellIndex);
      CellStyle style = templateCell.getCellStyle();
      if (style != null) {
        targetCell.setCellStyle(style);
      }
    }
  }

  /**
   * 为改单元格应用红色字体样式。
   *
   * @param workbook 当前工作簿。
   * @param cell 目标单元格。
   */
  private void applyRedStyle(Workbook workbook, Cell cell) {
    // 业务上只改字体颜色，不改边框、底色、对齐和合并布局，确保原始模板样式仍然成立。
    CellStyle originalStyle = cell.getCellStyle();
    String cacheKey = buildStyleCacheKey(cell);
    CellStyle redStyle = redStyleCache.get(cacheKey);
    if (redStyle == null) {
      redStyle = workbook.createCellStyle();
      if (originalStyle != null) {
        redStyle.cloneStyleFrom(originalStyle);
      }
      Font redFont = workbook.createFont();
      if (originalStyle != null) {
        Font originalFont = workbook.getFontAt(originalStyle.getFontIndex());
        redFont.setBold(originalFont.getBold());
        redFont.setFontHeight(originalFont.getFontHeight());
        redFont.setFontName(originalFont.getFontName());
        redFont.setItalic(originalFont.getItalic());
        redFont.setStrikeout(originalFont.getStrikeout());
        redFont.setTypeOffset(originalFont.getTypeOffset());
        redFont.setUnderline(originalFont.getUnderline());
        redFont.setCharSet(originalFont.getCharSet());
      }
      redFont.setColor(IndexedColors.RED.getIndex());
      redStyle.setFont(redFont);
      if (redStyle.getFillPattern() == FillPatternType.NO_FILL && originalStyle != null) {
        redStyle.setFillPattern(originalStyle.getFillPattern());
      }
      redStyleCache.put(cacheKey, redStyle);
    }
    cell.setCellStyle(redStyle);

    // 业务上字符串格再额外走一次 RichText，避免部分模板在字体索引切换后显示不刷新。
    if (cell.getCellType() == CellType.STRING) {
      RichTextString richText = workbook.getCreationHelper().createRichTextString(cell.getStringCellValue());
      richText.applyFont(workbook.getFontAt(redStyle.getFontIndex()));
      cell.setCellValue(richText);
    }
  }

  /**
   * 生成样式缓存键。
   *
   * @param cell 目标单元格。
   * @return 缓存键。
   */
  private String buildStyleCacheKey(Cell cell) {
    // 业务上同一工作簿里同样原样式的改单元格可以复用同一份红色派生样式。
    int styleIndex = cell.getCellStyle() == null ? -1 : cell.getCellStyle().getIndex();
    return cell.getSheet().getSheetName() + "#" + styleIndex;
  }

  /**
   * 读取单元格原值。
   *
   * @param cell 目标单元格。
   * @return 原值字符串。
   */
  private String readCellText(Cell cell) {
    // 业务上改单元格前先把旧值记入报告，便于后续人工复核修正是否准确。
    if (cell == null) {
      return "";
    }
    return switch (cell.getCellType()) {
      case STRING -> cell.getStringCellValue();
      case NUMERIC -> Double.toString(cell.getNumericCellValue());
      case BOOLEAN -> Boolean.toString(cell.getBooleanCellValue());
      case FORMULA -> cell.getCellFormula();
      case BLANK -> "";
      default -> cell.toString();
    };
  }

  /**
   * 获取或创建目标单元格。
   *
   * @param sheet 目标页签。
   * @param cellRef 单元格坐标。
   * @return 对应单元格。
   */
  private Cell getOrCreateCell(Sheet sheet, String cellRef) {
    // 业务上所有坐标都以 Excel 绝对坐标声明，避免列号计算出错。
    CellReference reference = new CellReference(cellRef);
    Row row = sheet.getRow(reference.getRow());
    if (row == null) {
      row = sheet.createRow(reference.getRow());
    }
    Cell cell = row.getCell(reference.getCol());
    if (cell == null) {
      cell = row.createCell(reference.getCol());
    }
    return cell;
  }

  /**
   * 输出 Markdown 改动报告。
   *
   * @throws IOException 写文件失败时抛出。
   */
  private void writeReport() throws IOException {
    // 业务上报告要落在 OPTION，便于和生成的新工作簿一起交付给后续复核者。
    try (BufferedWriter writer = Files.newBufferedWriter(REPORT_PATH, StandardCharsets.UTF_8)) {
      writer.write("# SBMAB203 文档修正报告\n\n");
      writer.write("- 输出目录: `" + OUTPUT_DOC_DIR + "`\n");
      writer.write("- 生成时间: 本次离线执行\n\n");

      for (WorkbookResult result : workbookResults) {
        writer.write("## " + result.sourceFileName() + "\n\n");
        if (!result.outputPath().isEmpty()) {
          writer.write("- 输出文件: `" + result.outputPath() + "`\n");
        }
        if (!result.summary().isEmpty()) {
          writer.write("- 核对结果: " + result.summary() + "\n\n");
          continue;
        }
        writer.write("- 改动数: " + result.changeRecords().size() + "\n\n");
        writer.write("| 页签 | 单元格 | 原值 | 新值 | 原因 |\n");
        writer.write("| --- | --- | --- | --- | --- |\n");
        for (ChangeRecord record : result.changeRecords()) {
          writer.write("| " + escapeMarkdown(record.sheetName()) + " | " + escapeMarkdown(record.cellRef()) + " | "
              + escapeMarkdown(record.oldValue()) + " | " + escapeMarkdown(record.newValue()) + " | "
              + escapeMarkdown(record.reason()) + " |\n");
        }
        writer.write("\n");
      }
    }
  }

  /**
   * 处理 Markdown 表格中的竖线与换行。
   *
   * @param value 原文本。
   * @return 可写入表格的文本。
   */
  private String escapeMarkdown(String value) {
    // 业务上改单元格说明里会出现 SQL 片段和条件句，需要先转义再进 Markdown 表格。
    if (value == null || value.isEmpty()) {
      return "";
    }
    return value.replace("|", "\\|").replace("\r", " ").replace("\n", "<br>");
  }

  /**
   * 工作簿操作函数式接口。
   */
  @FunctionalInterface
  private interface WorkbookOperation {

    /**
     * 对工作簿执行一次修正。
     *
     * @param workbook 当前工作簿。
     * @param records 改动记录列表。
     */
    void apply(Workbook workbook, List<ChangeRecord> records);
  }

  /**
   * 单个工作簿的计划。
   *
   * @param sourceFileName 原始文件名。
   * @param outputFileName 输出文件名。
   * @param operations 执行步骤。
   */
  private record WorkbookPlan(String sourceFileName, String outputFileName, List<WorkbookOperation> operations) {
  }

  /**
   * 单条改单元格记录。
   *
   * @param sheetName 页签名。
   * @param cellRef 单元格坐标。
   * @param oldValue 原值。
   * @param newValue 新值。
   * @param reason 改动原因。
   */
  private record ChangeRecord(String sheetName, String cellRef, String oldValue, String newValue, String reason) {
  }

  /**
   * 单个工作簿的执行结果。
   *
   * @param sourceFileName 原始文件名。
   * @param outputPath 输出路径。
   * @param changeRecords 改动记录。
   * @param summary 核对总结。
   */
  private record WorkbookResult(
      String sourceFileName,
      String outputPath,
      List<ChangeRecord> changeRecords,
      String summary) {
  }
}
