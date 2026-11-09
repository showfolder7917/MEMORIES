package jp.or.jasdec.sbf.cp.ma.batch.biz.sub;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.function.Predicate;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.CRC32;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * DB 数据生成工具（xlsm）与 testCase JSON 的通用同步工具。
 *
 * <p>本工具不使用 POI 保存 Workbook。它保留原始 ZIP 部件和 worksheet 外壳，只替换已有
 * input/expect 数据页的 {@code <sheetData>}。因此宏、ActiveX、关系、Office 扩展命名空间和
 * {@code mc:Ignorable} 不会被通用 XML 序列化器重写。</p>
 *
 * <p>通用性来自 define JSON 和工作表 C4 的物理表名，不包含任何业务表专用分支。</p>
 */
public final class DbXlsmImportTool {

  private static final ObjectMapper OBJECT_MAPPER = new ObjectMapper();
  private static final String WORKSHEET_NAMESPACE =
      "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  private static final String OFFICE_REL_NAMESPACE =
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  private static final int DATA_START_COLUMN = 11;
  private static final int OUTPUT_NUMBER_ROW = 9;
  private static final int TEST_CASE_ID_ROW = 10;
  private static final int EXCLUDED_ROW = 11;
  private static final int OPERATION_ROW = 12;
  private static final int HEADER_ROW = 13;
  private static final int DATA_START_ROW = 14;
  private static final Pattern SHEET_DATA_PATTERN = Pattern.compile(
      "<sheetData(?:\\s[^>]*)?>.*?</sheetData>", Pattern.DOTALL);
  private static final Pattern ROW_PATTERN = Pattern.compile(
      "<row\\b[^>]*\\br=\"(\\d+)\"[^>]*>.*?</row>", Pattern.DOTALL);
  private static final Pattern CELL_REFERENCE_PATTERN = Pattern.compile(
      "\\br=\"([A-Z]+)(\\d+)\"");
  private static final Pattern STYLE_PATTERN = Pattern.compile("\\bs=\"([^\"]+)\"");

  private DbXlsmImportTool() {
    // Utility class.
  }

  /**
   * 无参数时自动发现当前工程的目标 xlsm 与 testCase 根目录；也可显式传入两者。
   *
   * @param args {@code <xlsmPath> <testCaseRootOrDir> [caseId...]}
   * @throws Exception 目录、JSON、OOXML 或写入校验失败时抛出
   */
  public static void main(String[] args) throws Exception {
    Invocation invocation = resolveInvocation(args);
    ImportResult result = importCases(invocation.xlsmPath(), invocation.caseDirectories());
    System.out.println("synced workbook: " + invocation.xlsmPath());
    System.out.println("backup: " + result.backupPath());
    System.out.println("cases=" + result.caseCount()
        + ", sheets=" + result.updatedSheetCount()
        + ", records=" + result.recordCount());
    System.out.println("vbaProject.bin SHA-256=" + result.vbaSha256());
  }

  /**
   * 单独 Case 的快照导入入口。导入后的数据页只包含该 Case。
   *
   * @param xlsmPath 目标工作簿
   * @param testCaseDir 单独 Case 目录
   * @throws IOException 导入失败时抛出
   */
  public static void importTestCase(Path xlsmPath, Path testCaseDir) throws IOException {
    importCases(xlsmPath, List.of(testCaseDir));
  }

  private static Invocation resolveInvocation(String[] args) throws IOException {
    if (args.length == 1) {
      throw new IllegalArgumentException(
          "Usage: DbXlsmImportTool <xlsmPath> <testCaseRootOrDir> [caseId...]");
    }
    if (args.length >= 2) {
      Path xlsmPath = Paths.get(args[0]).toAbsolutePath().normalize();
      Path sourcePath = Paths.get(args[1]).toAbsolutePath().normalize();
      return new Invocation(xlsmPath, resolveCaseDirectories(sourcePath, args));
    }

    Path projectRoot = Paths.get("").toAbsolutePath().normalize();
    Path testDirectory = projectRoot.resolve("test");
    Path xlsmPath = findSinglePath(testDirectory, 1,
        path -> path.getFileName().toString().startsWith("020_DB")
            && path.getFileName().toString().endsWith("v2.15_template_CP.xlsm")
            && !path.getFileName().toString().contains(".before"),
        "target xlsm");
    Path testCaseRoot = findSinglePath(testDirectory.resolve("inputData"), 8,
        path -> Files.isDirectory(path)
            && "testCase".equals(path.getFileName().toString())
            && path.getParent() != null
            && "db".equals(path.getParent().getFileName().toString()),
        "db/testCase directory");
    return new Invocation(xlsmPath, resolveCaseDirectories(testCaseRoot, new String[0]));
  }

  private static Path findSinglePath(
      Path root, int maxDepth, Predicate<Path> predicate, String description) throws IOException {
    if (!Files.isDirectory(root)) {
      throw new IOException("Directory not found while discovering " + description + ": " + root);
    }
    try (var stream = Files.find(root, maxDepth, (path, attributes) -> predicate.test(path))) {
      List<Path> matches = stream.sorted().toList();
      if (matches.size() != 1) {
        throw new IOException("Expected exactly one " + description + " under " + root
            + ", but found " + matches);
      }
      return matches.get(0).toAbsolutePath().normalize();
    }
  }

  private static List<Path> resolveCaseDirectories(Path sourcePath, String[] args) throws IOException {
    if (args.length >= 3) {
      List<Path> selected = new ArrayList<>();
      for (int index = 2; index < args.length; index++) {
        Path caseDirectory = sourcePath.resolve(normalizeCaseId(args[index]));
        requireCaseDirectory(caseDirectory);
        selected.add(caseDirectory);
      }
      return selected.stream().sorted(Comparator.comparing(DbXlsmImportTool::caseId)).toList();
    }
    if (isCaseDirectory(sourcePath)) {
      return List.of(sourcePath);
    }
    if (!Files.isDirectory(sourcePath)) {
      throw new IOException("testCase root not found: " + sourcePath);
    }
    try (var stream = Files.list(sourcePath)) {
      List<Path> cases = stream.filter(Files::isDirectory)
          .filter(DbXlsmImportTool::isCaseDirectory)
          .sorted(Comparator.comparing(DbXlsmImportTool::caseId))
          .toList();
      if (cases.isEmpty()) {
        throw new IOException("No testCase directory found under: " + sourcePath);
      }
      return cases;
    }
  }

  private static boolean isCaseDirectory(Path path) {
    return Files.isDirectory(path.resolve("input")) && Files.isDirectory(path.resolve("expect"));
  }

  private static void requireCaseDirectory(Path path) throws IOException {
    if (!isCaseDirectory(path)) {
      throw new IOException("Invalid testCase directory: " + path);
    }
  }

  private static String caseId(Path caseDirectory) {
    return normalizeCaseId(caseDirectory.getFileName().toString());
  }

  private static String normalizeCaseId(String rawCaseId) {
    String normalized = rawCaseId == null ? "" : rawCaseId.trim();
    if (normalized.startsWith("testCase")) {
      normalized = normalized.substring("testCase".length());
    }
    if (!normalized.matches("\\d{6}")) {
      throw new IllegalArgumentException("testCaseId must be 6 digits: " + rawCaseId);
    }
    return normalized;
  }

  private static ImportResult importCases(Path xlsmPath, List<Path> caseDirectories) throws IOException {
    if (!Files.isRegularFile(xlsmPath)) {
      throw new IOException("xlsm not found: " + xlsmPath);
    }
    if (caseDirectories.isEmpty()) {
      throw new IOException("No test cases selected.");
    }

    Path caseRoot = caseDirectories.get(0).getParent();
    for (Path caseDirectory : caseDirectories) {
      requireCaseDirectory(caseDirectory);
      if (!caseRoot.equals(caseDirectory.getParent())) {
        throw new IOException("All test cases must have the same root: " + caseDirectories);
      }
    }
    Path defineDirectory = caseRoot.getParent().resolve("define");
    Map<String, List<String>> definitions = readDefinitions(defineDirectory);
    LinkedHashMap<String, ZipPart> parts = readZip(xlsmPath);
    byte[] macroBefore = requiredPart(parts, "xl/vbaProject.bin").bytes();
    List<String> sharedStrings = readSharedStrings(parts.get("xl/sharedStrings.xml"));
    Set<String> registeredCaseIds = readRegisteredCaseIds(parts, sharedStrings);
    List<Path> activeCaseDirectories = caseDirectories.stream()
        .filter(path -> registeredCaseIds.contains(caseId(path)))
        .toList();
    if (activeCaseDirectories.isEmpty()) {
      throw new IOException("No selected test case is registered in テストケース一覧: " + caseDirectories);
    }
    Map<SheetKey, String> dataSheets = discoverDataSheets(parts, sharedStrings);

    int updatedSheetCount = 0;
    int recordCount = 0;
    Set<String> updatedParts = new LinkedHashSet<>();
    for (Map.Entry<SheetKey, String> entry : dataSheets.entrySet()) {
      SheetKey key = entry.getKey();
      List<String> definedFields = definitions.get(key.physicalTableName());
      if (definedFields == null) {
        throw new IOException("define JSON not found for sheet " + key + ".");
      }
      String partName = entry.getValue();
      ZipPart originalPart = requiredPart(parts, partName);
      List<String> worksheetFields = readWorksheetFields(originalPart.bytes(), sharedStrings);
      if (!worksheetFields.equals(definedFields)) {
        throw new IOException("Worksheet definition differs from define JSON: " + key
            + "\nworksheet=" + worksheetFields + "\ndefine=" + definedFields);
      }

      List<CaseRecord> records = readCaseRecords(activeCaseDirectories, key, definedFields);
      byte[] rebuilt = rebuildSheetData(originalPart.bytes(), records, definedFields);
      parts.put(partName, originalPart.withBytes(rebuilt));
      updatedParts.add(partName);
      updatedSheetCount++;
      recordCount += records.size();
    }
    validateAllJsonTablesHaveSheets(activeCaseDirectories, dataSheets.keySet());

    Path backupPath = backupPath(xlsmPath);
    Files.copy(xlsmPath, backupPath, StandardCopyOption.REPLACE_EXISTING);
    Path temporary = Files.createTempFile(xlsmPath.getParent(), "db-xlsm-import-", ".xlsm");
    try {
      writeZip(temporary, parts);
      validateWrittenWorkbook(temporary, parts.keySet(), updatedParts, macroBefore);
      Files.move(temporary, xlsmPath, StandardCopyOption.REPLACE_EXISTING);
    } finally {
      Files.deleteIfExists(temporary);
    }
    return new ImportResult(
        backupPath, activeCaseDirectories.size(), updatedSheetCount, recordCount, sha256Hex(macroBefore));
  }

  private static Set<String> readRegisteredCaseIds(
      Map<String, ZipPart> parts, List<String> sharedStrings) throws IOException {
    Map<String, String> sheetParts = discoverSheetParts(parts);
    String listPartName = sheetParts.get("テストケース一覧");
    if (listPartName == null) {
      throw new IOException("テストケース一覧 sheet not found.");
    }
    Document document = parseXml(requiredPart(parts, listPartName).bytes());
    Set<String> result = new LinkedHashSet<>();
    NodeList cells = document.getElementsByTagNameNS(WORKSHEET_NAMESPACE, "c");
    for (int index = 0; index < cells.getLength(); index++) {
      Element cell = (Element) cells.item(index);
      Matcher matcher = Pattern.compile("B(\\d+)").matcher(cell.getAttribute("r"));
      if (!matcher.matches() || Integer.parseInt(matcher.group(1)) < 7) {
        continue;
      }
      String value = elementCellValue(cell, sharedStrings).trim();
      if (value.startsWith("testCase")) {
        value = value.substring("testCase".length());
      }
      if (value.matches("\\d{6}")) {
        result.add(value);
      }
    }
    if (result.isEmpty()) {
      throw new IOException("No testCaseId found in テストケース一覧.");
    }
    return result;
  }

  private static Map<String, List<String>> readDefinitions(Path defineDirectory) throws IOException {
    if (!Files.isDirectory(defineDirectory)) {
      throw new IOException("define directory not found: " + defineDirectory);
    }
    Map<String, List<String>> definitions = new HashMap<>();
    try (var stream = Files.list(defineDirectory)) {
      for (Path jsonFile : stream.filter(path -> path.toString().endsWith(".json")).sorted().toList()) {
        Map<String, Object> json = readJsonObject(jsonFile);
        String table = String.valueOf(json.getOrDefault("physicalTableName", ""));
        List<String> fields = new ArrayList<>();
        Object structure = json.get("tableStructure");
        if (structure instanceof List<?> fieldList) {
          for (Object rawField : fieldList) {
            if (rawField instanceof Map<?, ?> field) {
              fields.add(String.valueOf(field.get("physicalName")));
            }
          }
        }
        if (!table.isBlank() && !fields.isEmpty()) {
          definitions.put(table, List.copyOf(fields));
        }
      }
    }
    return definitions;
  }

  private static Map<String, Object> readJsonObject(Path jsonFile) throws IOException {
    return OBJECT_MAPPER.readValue(Files.readString(jsonFile, StandardCharsets.UTF_8),
        new TypeReference<LinkedHashMap<String, Object>>() {});
  }

  private static List<CaseRecord> readCaseRecords(
      List<Path> caseDirectories, SheetKey key, List<String> fields) throws IOException {
    List<CaseRecord> records = new ArrayList<>();
    String directoryName = "inp".equals(key.mode()) ? "input" : "expect";
    for (Path caseDirectory : caseDirectories) {
      Path jsonFile = caseDirectory.resolve(directoryName).resolve(key.physicalTableName() + ".json");
      if (!Files.isRegularFile(jsonFile)) {
        continue;
      }
      Map<String, Object> payload = readJsonObject(jsonFile);
      String payloadTable = String.valueOf(payload.getOrDefault("physicalTableName", ""));
      if (!key.physicalTableName().equals(payloadTable)) {
        throw new IOException("physicalTableName mismatch: " + jsonFile);
      }
      Object rawTableData = payload.get("tableData");
      if (!(rawTableData instanceof List<?> tableData)) {
        throw new IOException("tableData must be an array: " + jsonFile);
      }
      for (int rowIndex = 0; rowIndex < tableData.size(); rowIndex++) {
        Object rawRecord = tableData.get(rowIndex);
        if (!(rawRecord instanceof Map<?, ?> sourceRecord)) {
          throw new IOException("tableData record must be an object: " + jsonFile + "[" + rowIndex + "]");
        }
        LinkedHashMap<String, Object> record = new LinkedHashMap<>();
        for (Map.Entry<?, ?> value : sourceRecord.entrySet()) {
          record.put(String.valueOf(value.getKey()), value.getValue());
        }
        if (!new ArrayList<>(record.keySet()).equals(fields)) {
          throw new IOException("JSON field order/set differs from define JSON: " + jsonFile
              + "[" + rowIndex + "]\njson=" + record.keySet() + "\ndefine=" + fields);
        }
        records.add(new CaseRecord(caseId(caseDirectory), record));
      }
    }
    return records;
  }

  private static void validateAllJsonTablesHaveSheets(
      List<Path> caseDirectories, Set<SheetKey> sheetKeys) throws IOException {
    Set<SheetKey> available = Set.copyOf(sheetKeys);
    for (Path caseDirectory : caseDirectories) {
      for (String directoryName : List.of("input", "expect")) {
        String mode = "input".equals(directoryName) ? "inp" : "exp";
        try (var stream = Files.list(caseDirectory.resolve(directoryName))) {
          for (Path jsonFile : stream.filter(path -> path.toString().endsWith(".json")).toList()) {
            String table = String.valueOf(readJsonObject(jsonFile).getOrDefault("physicalTableName", ""));
            SheetKey key = new SheetKey(mode, table);
            if (!available.contains(key)) {
              throw new IOException("Workbook data sheet not found for " + key + ": " + jsonFile);
            }
          }
        }
      }
    }
  }

  private static Map<SheetKey, String> discoverDataSheets(
      Map<String, ZipPart> parts, List<String> sharedStrings) throws IOException {
    Map<String, String> sheetParts = discoverSheetParts(parts);
    Map<SheetKey, String> result = new LinkedHashMap<>();
    for (Map.Entry<String, String> sheetEntry : sheetParts.entrySet()) {
      String sheetName = sheetEntry.getKey();
      String mode = sheetName.startsWith("inp_") ? "inp" : sheetName.startsWith("exp_") ? "exp" : "";
      if (mode.isBlank()) {
        continue;
      }
      String partName = sheetEntry.getValue();
      ZipPart worksheetPart = requiredPart(parts, partName);
      String physicalTableName = readCellValue(worksheetPart.bytes(), "C4", sharedStrings);
      if (!physicalTableName.isBlank()) {
        SheetKey key = new SheetKey(mode, physicalTableName);
        if (result.put(key, partName) != null) {
          throw new IOException("Duplicate data sheet for " + key);
        }
      }
    }
    return result;
  }

  private static Map<String, String> discoverSheetParts(Map<String, ZipPart> parts) throws IOException {
    Document workbook = parseXml(requiredPart(parts, "xl/workbook.xml").bytes());
    Document relationships = parseXml(requiredPart(parts, "xl/_rels/workbook.xml.rels").bytes());
    Map<String, String> targets = new HashMap<>();
    NodeList relationshipNodes = relationships.getElementsByTagNameNS("*", "Relationship");
    for (int index = 0; index < relationshipNodes.getLength(); index++) {
      Element relation = (Element) relationshipNodes.item(index);
      targets.put(relation.getAttribute("Id"), normalizePartName("xl", relation.getAttribute("Target")));
    }

    Map<String, String> result = new LinkedHashMap<>();
    NodeList sheets = workbook.getElementsByTagNameNS(WORKSHEET_NAMESPACE, "sheet");
    for (int index = 0; index < sheets.getLength(); index++) {
      Element sheet = (Element) sheets.item(index);
      String sheetName = sheet.getAttribute("name");
      String relationId = sheet.getAttributeNS(OFFICE_REL_NAMESPACE, "id");
      String partName = targets.get(relationId);
      if (partName == null || !parts.containsKey(partName)) {
        throw new IOException("Worksheet relationship target missing: " + sheetName + " -> " + partName);
      }
      result.put(sheetName, partName);
    }
    return result;
  }

  private static String normalizePartName(String base, String target) {
    if (target.startsWith("/")) {
      return target.substring(1);
    }
    Path normalized = Paths.get(base).resolve(target.replace('/', java.io.File.separatorChar)).normalize();
    return normalized.toString().replace(java.io.File.separatorChar, '/');
  }

  private static List<String> readSharedStrings(ZipPart part) throws IOException {
    if (part == null) {
      return List.of();
    }
    Document document = parseXml(part.bytes());
    NodeList items = document.getElementsByTagNameNS(WORKSHEET_NAMESPACE, "si");
    List<String> result = new ArrayList<>();
    for (int index = 0; index < items.getLength(); index++) {
      result.add(descendantText((Element) items.item(index), "t"));
    }
    return result;
  }

  private static String readCellValue(byte[] worksheet, String reference, List<String> sharedStrings)
      throws IOException {
    Document document = parseXml(worksheet);
    NodeList cells = document.getElementsByTagNameNS(WORKSHEET_NAMESPACE, "c");
    for (int index = 0; index < cells.getLength(); index++) {
      Element cell = (Element) cells.item(index);
      if (reference.equals(cell.getAttribute("r"))) {
        return elementCellValue(cell, sharedStrings);
      }
    }
    return "";
  }

  private static List<String> readWorksheetFields(byte[] worksheet, List<String> sharedStrings)
      throws IOException {
    Document document = parseXml(worksheet);
    TreeMap<Integer, String> fields = new TreeMap<>();
    NodeList cells = document.getElementsByTagNameNS(WORKSHEET_NAMESPACE, "c");
    for (int index = 0; index < cells.getLength(); index++) {
      Element cell = (Element) cells.item(index);
      Matcher matcher = Pattern.compile("D(\\d+)").matcher(cell.getAttribute("r"));
      if (matcher.matches()) {
        int row = Integer.parseInt(matcher.group(1));
        String value = elementCellValue(cell, sharedStrings);
        if (row >= DATA_START_ROW && !value.isBlank()) {
          fields.put(row, value);
        }
      }
    }
    List<String> result = new ArrayList<>();
    int expectedRow = DATA_START_ROW;
    for (Map.Entry<Integer, String> field : fields.entrySet()) {
      if (field.getKey() != expectedRow++) {
        throw new IOException("Non-contiguous worksheet field rows: " + fields.keySet());
      }
      result.add(field.getValue());
    }
    return result;
  }

  private static String elementCellValue(Element cell, List<String> sharedStrings) throws IOException {
    String type = cell.getAttribute("t");
    if ("inlineStr".equals(type)) {
      return descendantText(cell, "t");
    }
    String value = descendantText(cell, "v");
    if ("s".equals(type) && !value.isBlank()) {
      int index = Integer.parseInt(value);
      if (index < 0 || index >= sharedStrings.size()) {
        throw new IOException("Invalid shared string index: " + index);
      }
      return sharedStrings.get(index);
    }
    return value;
  }

  private static String descendantText(Element parent, String localName) {
    NodeList nodes = parent.getElementsByTagNameNS(WORKSHEET_NAMESPACE, localName);
    StringBuilder result = new StringBuilder();
    for (int index = 0; index < nodes.getLength(); index++) {
      result.append(nodes.item(index).getTextContent());
    }
    return result.toString();
  }

  private static byte[] rebuildSheetData(
      byte[] worksheetBytes, List<CaseRecord> records, List<String> fields) throws IOException {
    String worksheet = new String(worksheetBytes, StandardCharsets.UTF_8);
    Matcher sheetDataMatcher = SHEET_DATA_PATTERN.matcher(worksheet);
    if (!sheetDataMatcher.find()) {
      throw new IOException("sheetData not found in worksheet.");
    }
    String originalSheetData = sheetDataMatcher.group();
    Matcher rowMatcher = ROW_PATTERN.matcher(originalSheetData);
    StringBuffer rebuiltSheetData = new StringBuffer();
    Set<Integer> requiredRows = new LinkedHashSet<>(
        List.of(OUTPUT_NUMBER_ROW, TEST_CASE_ID_ROW, EXCLUDED_ROW, OPERATION_ROW, HEADER_ROW));
    for (int index = 0; index < fields.size(); index++) {
      requiredRows.add(DATA_START_ROW + index);
    }
    Set<Integer> foundRows = new LinkedHashSet<>();
    while (rowMatcher.find()) {
      int rowNumber = Integer.parseInt(rowMatcher.group(1));
      String originalRow = rowMatcher.group();
      String style = templateStyle(originalRow);
      String rowWithoutData = removeDataCells(originalRow);
      if (requiredRows.contains(rowNumber)) {
        foundRows.add(rowNumber);
        StringBuilder cells = new StringBuilder();
        for (int recordIndex = 0; recordIndex < records.size(); recordIndex++) {
          String reference = columnName(DATA_START_COLUMN + recordIndex) + rowNumber;
          Object value;
          if (rowNumber == OUTPUT_NUMBER_ROW) {
            value = recordIndex + 1;
          } else if (rowNumber == TEST_CASE_ID_ROW) {
            value = records.get(recordIndex).caseId();
          } else if (rowNumber == EXCLUDED_ROW || rowNumber == OPERATION_ROW || rowNumber == HEADER_ROW) {
            value = null;
          } else {
            String field = fields.get(rowNumber - DATA_START_ROW);
            value = records.get(recordIndex).record().get(field);
          }
          cells.append(cellXml(reference, style, value));
        }
        rowWithoutData = rowWithoutData.replace("</row>", cells + "</row>");
      }
      rowMatcher.appendReplacement(rebuiltSheetData, Matcher.quoteReplacement(rowWithoutData));
    }
    rowMatcher.appendTail(rebuiltSheetData);
    if (!foundRows.equals(requiredRows)) {
      Set<Integer> missing = new LinkedHashSet<>(requiredRows);
      missing.removeAll(foundRows);
      throw new IOException("Required worksheet rows are missing: " + missing);
    }
    String result = worksheet.substring(0, sheetDataMatcher.start())
        + rebuiltSheetData + worksheet.substring(sheetDataMatcher.end());
    validateWorksheetNamespaces(result);
    byte[] rebuiltBytes = result.getBytes(StandardCharsets.UTF_8);
    validateCellStructure(rebuiltBytes);
    return rebuiltBytes;
  }

  private static String removeDataCells(String rowXml) throws IOException {
    StringBuilder result = new StringBuilder(rowXml.length());
    int cursor = 0;
    for (CellFragment cell : cellsInRow(rowXml)) {
      result.append(rowXml, cursor, cell.start());
      if (cell.column() < DATA_START_COLUMN) {
        result.append(cell.xml());
      }
      cursor = cell.end();
    }
    result.append(rowXml, cursor, rowXml.length());
    return result.toString();
  }

  private static String templateStyle(String rowXml) throws IOException {
    String fallback = "";
    for (CellFragment cell : cellsInRow(rowXml)) {
      Matcher styleMatcher = STYLE_PATTERN.matcher(cell.xml());
      String style = styleMatcher.find() ? styleMatcher.group(1) : "";
      if (cell.column() == DATA_START_COLUMN) {
        return style;
      }
      if (cell.column() >= DATA_START_COLUMN && fallback.isBlank()) {
        fallback = style;
      }
    }
    return fallback;
  }

  private static List<CellFragment> cellsInRow(String rowXml) throws IOException {
    List<CellFragment> cells = new ArrayList<>();
    int searchFrom = 0;
    while (true) {
      int start = rowXml.indexOf("<c ", searchFrom);
      if (start < 0) {
        return cells;
      }
      int openingEnd = rowXml.indexOf('>', start);
      if (openingEnd < 0) {
        throw new IOException("Cell opening tag is incomplete.");
      }
      boolean selfClosing = rowXml.charAt(openingEnd - 1) == '/';
      int end;
      if (selfClosing) {
        end = openingEnd + 1;
      } else {
        int closingStart = rowXml.indexOf("</c>", openingEnd + 1);
        if (closingStart < 0) {
          throw new IOException("Cell closing tag is missing.");
        }
        end = closingStart + "</c>".length();
      }
      String cellXml = rowXml.substring(start, end);
      Matcher referenceMatcher = CELL_REFERENCE_PATTERN.matcher(cellXml);
      if (!referenceMatcher.find()) {
        throw new IOException("Cell reference is missing: " + cellXml);
      }
      cells.add(new CellFragment(
          start, end, columnNumber(referenceMatcher.group(1)), cellXml));
      searchFrom = end;
    }
  }

  private static String cellXml(String reference, String style, Object value) {
    String styleAttribute = style.isBlank() ? "" : " s=\"" + xmlAttribute(style) + "\"";
    if (value == null) {
      return "<c r=\"" + reference + "\"" + styleAttribute + "/>";
    }
    if (value instanceof Boolean bool) {
      return "<c r=\"" + reference + "\"" + styleAttribute + " t=\"b\"><v>"
          + (bool ? "1" : "0") + "</v></c>";
    }
    if (value instanceof Number number) {
      return "<c r=\"" + reference + "\"" + styleAttribute + "><v>"
          + xmlText(number.toString()) + "</v></c>";
    }
    String text = String.valueOf(value);
    String preserve = !text.isEmpty()
        && (Character.isWhitespace(text.charAt(0)) || Character.isWhitespace(text.charAt(text.length() - 1)))
        ? " xml:space=\"preserve\"" : "";
    String textNode = text.isEmpty() ? "<t/>" : "<t" + preserve + ">" + xmlText(text) + "</t>";
    return "<c r=\"" + reference + "\"" + styleAttribute
        + " t=\"inlineStr\"><is>" + textNode + "</is></c>";
  }

  private static int columnNumber(String letters) {
    int result = 0;
    for (int index = 0; index < letters.length(); index++) {
      result = result * 26 + letters.charAt(index) - 'A' + 1;
    }
    return result;
  }

  private static String columnName(int number) {
    StringBuilder result = new StringBuilder();
    while (number > 0) {
      number--;
      result.insert(0, (char) ('A' + number % 26));
      number /= 26;
    }
    return result.toString();
  }

  private static String xmlText(String value) {
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
  }

  private static String xmlAttribute(String value) {
    return xmlText(value).replace("\"", "&quot;").replace("'", "&apos;");
  }

  private static void validateWorksheetNamespaces(String worksheet) throws IOException {
    Matcher rootMatcher = Pattern.compile("<worksheet\\b[^>]*>").matcher(worksheet);
    if (!rootMatcher.find()) {
      throw new IOException("worksheet root element not found.");
    }
    String root = rootMatcher.group();
    Matcher ignorableMatcher = Pattern.compile("\\bIgnorable=\"([^\"]+)\"").matcher(root);
    if (ignorableMatcher.find()) {
      for (String prefix : ignorableMatcher.group(1).split("\\s+")) {
        if (!root.contains("xmlns:" + prefix + "=")) {
          throw new IOException("mc:Ignorable prefix is not declared: " + prefix);
        }
      }
    }
    if (Pattern.compile("xmlns:ns\\d+=").matcher(root).find()) {
      throw new IOException("Auto-generated namespace prefix detected in worksheet root.");
    }
  }

  private static LinkedHashMap<String, ZipPart> readZip(Path xlsmPath) throws IOException {
    LinkedHashMap<String, ZipPart> result = new LinkedHashMap<>();
    try (ZipFile zip = new ZipFile(xlsmPath.toFile(), StandardCharsets.UTF_8)) {
      Enumeration<? extends ZipEntry> entries = zip.entries();
      while (entries.hasMoreElements()) {
        ZipEntry entry = entries.nextElement();
        try (InputStream input = zip.getInputStream(entry)) {
          byte[] bytes = input.readAllBytes();
          result.put(entry.getName(), new ZipPart(
              entry.getName(), entry.getMethod(), entry.getTime(), entry.getComment(), entry.getExtra(), bytes));
        }
      }
    }
    return result;
  }

  private static ZipPart requiredPart(Map<String, ZipPart> parts, String name) throws IOException {
    ZipPart part = parts.get(name);
    if (part == null) {
      throw new IOException("Required XLSM part not found: " + name);
    }
    return part;
  }

  private static void writeZip(Path target, LinkedHashMap<String, ZipPart> parts) throws IOException {
    try (ZipOutputStream output = new ZipOutputStream(Files.newOutputStream(target), StandardCharsets.UTF_8)) {
      for (ZipPart part : parts.values()) {
        ZipEntry entry = new ZipEntry(part.name());
        entry.setMethod(part.method());
        if (part.time() >= 0) {
          entry.setTime(part.time());
        }
        if (part.comment() != null) {
          entry.setComment(part.comment());
        }
        if (part.extra() != null) {
          entry.setExtra(part.extra());
        }
        if (part.method() == ZipEntry.STORED) {
          CRC32 crc = new CRC32();
          crc.update(part.bytes());
          entry.setSize(part.bytes().length);
          entry.setCompressedSize(part.bytes().length);
          entry.setCrc(crc.getValue());
        }
        output.putNextEntry(entry);
        output.write(part.bytes());
        output.closeEntry();
      }
    }
  }

  private static void validateWrittenWorkbook(
      Path workbook, Set<String> expectedParts, Set<String> updatedParts, byte[] macroBefore) throws IOException {
    Set<String> actualParts = new LinkedHashSet<>();
    byte[] macroAfter = null;
    try (ZipFile zip = new ZipFile(workbook.toFile(), StandardCharsets.UTF_8)) {
      Enumeration<? extends ZipEntry> entries = zip.entries();
      while (entries.hasMoreElements()) {
        ZipEntry entry = entries.nextElement();
        actualParts.add(entry.getName());
        try (InputStream input = zip.getInputStream(entry)) {
          byte[] bytes = input.readAllBytes();
          if (updatedParts.contains(entry.getName())) {
            parseXml(bytes);
            validateWorksheetNamespaces(new String(bytes, StandardCharsets.UTF_8));
            validateCellStructure(bytes);
          }
          if ("xl/vbaProject.bin".equals(entry.getName())) {
            macroAfter = bytes;
          }
        }
      }
    }
    if (!actualParts.equals(expectedParts)) {
      throw new IOException("ZIP part list changed unexpectedly.");
    }
    if (!Arrays.equals(macroBefore, macroAfter)) {
      throw new IOException("vbaProject.bin changed unexpectedly.");
    }
  }

  private static void validateCellStructure(byte[] worksheetBytes) throws IOException {
    Document document = parseXml(worksheetBytes);
    NodeList rows = document.getElementsByTagNameNS(WORKSHEET_NAMESPACE, "row");
    Set<String> workbookReferences = new LinkedHashSet<>();
    for (int rowIndex = 0; rowIndex < rows.getLength(); rowIndex++) {
      Element row = (Element) rows.item(rowIndex);
      int rowNumber = Integer.parseInt(row.getAttribute("r"));
      int previousColumn = 0;
      NodeList cells = row.getElementsByTagNameNS(WORKSHEET_NAMESPACE, "c");
      for (int cellIndex = 0; cellIndex < cells.getLength(); cellIndex++) {
        Element cell = (Element) cells.item(cellIndex);
        String reference = cell.getAttribute("r");
        Matcher referenceMatcher = CELL_REFERENCE_PATTERN.matcher("r=\"" + reference + "\"");
        if (!referenceMatcher.find()) {
          throw new IOException("Invalid cell reference: " + reference);
        }
        int column = columnNumber(referenceMatcher.group(1));
        int referencedRow = Integer.parseInt(referenceMatcher.group(2));
        if (referencedRow != rowNumber) {
          throw new IOException("Cell row does not match parent row: " + reference);
        }
        if (!workbookReferences.add(reference)) {
          throw new IOException("Duplicate cell reference: " + reference);
        }
        if (column <= previousColumn) {
          throw new IOException("Cell references are not strictly ordered at: " + reference);
        }
        previousColumn = column;
      }
    }
  }

  private static Document parseXml(byte[] bytes) throws IOException {
    try {
      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
      factory.setNamespaceAware(true);
      factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
      factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
      factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
      factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, "");
      factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_SCHEMA, "");
      return factory.newDocumentBuilder().parse(new ByteArrayInputStream(bytes));
    } catch (ParserConfigurationException | SAXException exception) {
      throw new IOException("Invalid OOXML part.", exception);
    }
  }

  private static Path backupPath(Path xlsmPath) {
    String fileName = xlsmPath.getFileName().toString();
    String stem = fileName.toLowerCase(Locale.ROOT).endsWith(".xlsm")
        ? fileName.substring(0, fileName.length() - 5) : fileName;
    return xlsmPath.resolveSibling(stem + ".before-db-import.xlsm");
  }

  private static String sha256Hex(byte[] bytes) throws IOException {
    try {
      byte[] digest = MessageDigest.getInstance("SHA-256").digest(bytes);
      StringBuilder result = new StringBuilder();
      for (byte value : digest) {
        result.append(String.format("%02x", value));
      }
      return result.toString();
    } catch (NoSuchAlgorithmException exception) {
      throw new IOException("SHA-256 is unavailable.", exception);
    }
  }

  private record Invocation(Path xlsmPath, List<Path> caseDirectories) {
  }

  private record ImportResult(
      Path backupPath, int caseCount, int updatedSheetCount, int recordCount, String vbaSha256) {
  }

  private record SheetKey(String mode, String physicalTableName) {
  }

  private record CaseRecord(String caseId, LinkedHashMap<String, Object> record) {
  }

  private record CellFragment(int start, int end, int column, String xml) {
  }

  private record ZipPart(
      String name, int method, long time, String comment, byte[] extra, byte[] bytes) {

    private ZipPart withBytes(byte[] replacement) {
      return new ZipPart(name, method, time, comment, extra, replacement);
    }
  }
}
