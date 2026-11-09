////////////////////////////////////////////////////////////////////////////
//
// プログラム名：Sub
// ファイル名：Sub
// 作成者：automatic generator
// 作成日：2026/01/14
// 備考：
//
////////////////////////////////////////////////////////////////////////////

package jp.or.jasdec.sbf.cp.ma.batch.biz;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.io.StringReader;
import java.io.StringWriter;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.nio.file.CopyOption;
import java.nio.file.DirectoryStream;
import java.nio.file.FileVisitOption;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.BasicFileAttributes;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.EnumSet;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;



import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
// import ch.qos.logback.classic.joran.JoranConfigurator;
// import oracle.security.o3logon.a;


/**
 * 
 */
public class Sub {
  // import org.slf4j.Logger;
  // import org.slf4j.LoggerFactory;
  // private static final Logger logger = LoggerFactory.getLogger(SBMAB002.class);

  // implementation "org.apache.poi:poi:5.2.5"
  // implementation "org.apache.poi:poi-ooxml:5.2.5"

  // implementation 'com.github.jsqlparser:jsqlparser:4.6'
  // implementation 'org.junit.platform:junit-platform-launcher:1.9.1'


  public static Class<Sub> clazz = Sub.class;

  // 遷移用 行う不可
  public static B b = getB();

  public static Util u = (Util) getClass("Util");

  public static Oracle o = (Oracle) getClass("Oracle");

  public static Excel e = (Excel) getClass("Excel");

  public static Content c = (Content) getClass("Content");
  public static Template t = (Template) getClass("Template");

  public static JsonUtils JsonUtils = (JsonUtils) getClass("JsonUtils");

  public static Application a = (Application) getClass("Application");
  public static ApplicationSub as = (ApplicationSub) getClass("ApplicationSub");

  public static Test my = (Test) getClass("Test");

  /**
   * 主処理
   * 
   * <pre>
    implementation "org.apache.poi:poi:5.2.5"
    implementation "org.apache.poi:poi-ooxml:5.2.5"
    implementation "org.eclipse.platform:org.eclipse.swt.win32.win32.x86_64:3.133.0"
    implementation 'com.github.kwhat:jnativehook:2.2.2'
  
    implementation 'com.github.jsqlparser:jsqlparser:4.6'
   * </pre>
   *
   * @param args パラメータ
   * 
   * @throws Exception -
   * 
   */
  public static void main(String[] args) {
    String[] fileNames = {"JsonUtils"};
    as.testInit(fileNames);
    // my.test();
  }

  /**
   * 
   */
  public static B getB() {
    B b = (B) getMainInstance("B");
    return b;
  }

  /**
  * 
  */
  public static Object getClass(String className) {
    Object o = getMainInstance(className);
    return o;
  }

  /**
   * innerName B
   */
  public static Object getMainInstance(String innerName) {

    try {

      Object outer = clazz.getDeclaredConstructor().newInstance(); // clazz = TestApplication.class

      Class<?> innerCls = Class.forName(outer.getClass().getName() + "$" + innerName);
      java.lang.reflect.Constructor<?> ctor = innerCls.getDeclaredConstructor(outer.getClass());
      ctor.setAccessible(true);

      Object o = ctor.newInstance(outer);

      return o;

    } catch (Exception e) {
      e.printStackTrace();
    }
    return null;
  }

  /**
  * 
  */
  public class JsonUtils {

    private static final Logger LOG = LoggerFactory.getLogger(JsonUtils.class);

    // からのフィールドをむしする
    private static ObjectMapper OBJECT_MAPPER = new ObjectMapper();

    private static ObjectMapper OBJECT_MAPPER_EXT = new ObjectMapper();

    static {

      DateFormat df = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
      // からちとからのコレクションをむしす
      // OBJECT_MAPPER.setSerializationInclusion(JsonInclude.Include.NON_EMPTY);
      OBJECT_MAPPER.setDateFormat(df);

      // 値が NULL の場合、属性値は返却されません。
      OBJECT_MAPPER_EXT.setSerializationInclusion(JsonInclude.Include.NON_NULL);
    }

    /**
    * 
    */
    public static String toJson(Object object) {
      try {
        String json = OBJECT_MAPPER.writeValueAsString(object);
        return json;
      } catch (Exception e) {
        throw new RuntimeException("Json.format error:" + object, e);
      }
    }

    /**
     * 如果对象为Null,则该实体不在json中显示. 如果集合为空集合,返回"[]".
     */
    public String toJsonExt(Object object) {
      try {
        return OBJECT_MAPPER_EXT.writeValueAsString(object);
      } catch (IOException e) {
        LOG.warn("write to json string error:" + object, e);
      }
      return null;
    }

    /**
    * 
    */
    public <T> T fromJson(String jsonString, Class<T> clazz) {
      if (jsonString == null) {
        return null;
      }
      try {
        return OBJECT_MAPPER.readValue(jsonString, clazz);
      } catch (IOException e) {
        throw new RuntimeException("parse json string error:" + jsonString, e);
      }
    }

    /**
     * 
     */
    public <T> T fromJson(String jsonString, TypeReference<T> typeReference) {
      if (jsonString == null) {
        return null;
      }
      try {
        return OBJECT_MAPPER.readValue(jsonString, typeReference);
      } catch (IOException e) {
        throw new RuntimeException("parse json string error:" + jsonString, e);
      }
    }
  }

  /**
   * Util
   */
  public class Util {
    static Sub sub;
    static Test t;
    static Application a;

    /**
     * <pre>
     * プロジェクト名を取得する
     *  例:「C:\opt\PT\IT\ITMAB203\」のプロジェクト名は 
     * 「ITMAB203」
     * </pre>
     */
    public String getProjectName() {
      String path = u.getProjectDir("");
      String[] pns = path.split("\\\\");
      String projectName = pns[pns.length - 1];
      return projectName;
    }

    /**
     * <pre>
     * 現在ファイル「Sub.java」のパスを取得する
     * 例:「C:\opt\PT\IT\ITMAB203\src\test\java\jp\or\jasdec\sbf\it\ma\batch\biz\Sub.java」のパスは
     * 「C:\opt\PT\IT\ITMAB203\src\test\java\jp\or\jasdec\sbf\it\ma\batch\biz\」
     * </pre>
     */
    public String getCurrentFilePath() {
      String testPath = u.getProjectDir("src");
      List<File> list = u.findFile(testPath);
      for (File f : list) {
        String className = Sub.class.getSimpleName() + ".java";
        if (f.getName().equals(className)) {
          return f.getParent() + File.separator;
        }
      }
      return null;
    }


    /**
     * 例 C:\opt\PT\SBMAB002_PT
     */
    public String getProjectDir(String... dir) {
      String projectPath = System.getProperty("user.dir");
      return Paths.get(projectPath, dir).toString() + File.separator;
    }

    /**
     * String dir = u.getProjectFindFirstDir("src/test", "biz");
     * 
     * return C:\opt\PT\SBMAB002_PT\src\test\java\jp\or\jasdec\sbf\sb\ma\batch\biz\
     */
    public String getProjectFindFirstDir(String path, String dirName) {
      return getFindFirstDir(getProjectDir(path), dirName).getAbsolutePath() + File.separator;
    }

    /**
     * 
     * File dir = getFindFirstDir("C:\opt\PT\SBMAB002_PT\src\test\", "biz");
     * 
     * return C:\opt\PT\SBMAB002_PT\src\test\java\jp\or\jasdec\sbf\sb\ma\batch\biz\
     */

    public File getFindFirstDir(String path, String dirName) {
      Path base = Paths.get(path);

      try (Stream<Path> s = Files.walk(base)) {
        return s.filter(Files::isDirectory).filter(p -> p.getFileName() != null && p.getFileName().toString().equals(dirName))
            .map(Path::toAbsolutePath).map(Path::toFile).findFirst().orElse(null); // ← 不抛异常
      } catch (IOException e) {
        System.out.println("IO异常: " + e.getMessage());
      }

      System.out.println("目录未找到: " + dirName + " -> path: " + base.toAbsolutePath());
      return null;
    }


    /**
     * 単一ファイルを取得する
     */
    public File findFile(String filePath, String fileName) {
      List<File> list = u.findFile(filePath);
      for (File f : list) {
        if (f.getName().equals(fileName)) {
          return f;
        }
      }
      return null;
    }

    /**
     * 全てのファイルをリストに書き込みます。
     */
    public List<File> findFile(String path) {
      File baseDir = new File(path);
      List<File> fileList = new ArrayList<File>();
      findFile(baseDir, fileList);
      return fileList;
    }

    /**
     * 全てのファイルをリストに書き込みます。 为了去掉重复
     */
    public Map<File, String> findFileToMap(String path) {

      Map<String, File> map = new HashMap<String, File>();
      File baseDir = new File(path);
      List<File> fileList = new ArrayList<File>();
      findFile(baseDir, fileList);
      for (File f : fileList) {
        String baseName = f.getName().substring(0, f.getName().lastIndexOf("."));
        map.put(baseName, f);
      }

      Map<File, String> returnMap = new HashMap<File, String>();
      for (String key : map.keySet()) {
        returnMap.put(map.get(key), key);
      }
      return returnMap;
    }

    /**
     * 全てのファイルをリストに書き込みます。
     *
     * @param baseDir フォルダーパス
     * @param fileList 全部ファイル
     * @return fileList -
     */
    private List<File> findFile(File baseDir, List<File> fileList) {
      File[] files = baseDir.listFiles();
      for (File file : files) {
        if (file.isDirectory()) {
          findFile(file, fileList);
        } else {
          fileList.add(file);
        }
      }
      return fileList;
    }

    /**
     * 改行で分割、リストに格納する
     *
     * @param content ファイルパス
     * 
     * @return リスト
     */
    public List<String> readContentToList(String content) {
      List<String> list = new ArrayList<String>();
      try (BufferedReader reader = new BufferedReader(new StringReader(content))) {
        String line;
        while ((line = reader.readLine()) != null) {
          list.add(line);
        }
      } catch (IOException e) {
        e.printStackTrace();
      }
      return list;
    }

    /**
     * 改行で分割、リストに格納する
     *
     * @param filePath ファイルパス
     * 
     * @return リスト
     */
    public List<String> readFileToList(String filePath) {
      String content = readFile(filePath);
      List<String> list = readContentToList(content);
      return list;
    }

    /**
     *
     * @param filePath -
     * @return -
     */
    public String readFile(String filePath) {
      StringBuilder sb = new StringBuilder();
      try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(filePath), StandardCharsets.UTF_8))) {

        String line;
        while ((line = reader.readLine()) != null) {
          sb.append(line + "\n");
        }

      } catch (IOException e) {
        e.printStackTrace();
      }
      return sb.toString();
    }

    /**
     * @param content -
     * @param filePath -
     */
    public void outFile(String content, String filePath) {
      File file = new File(filePath);
      file.getParentFile().mkdirs();
      try (PrintWriter writer = new PrintWriter(new BufferedWriter(new OutputStreamWriter(new FileOutputStream(filePath), StandardCharsets.UTF_8)))) {
        writer.println(content);
      } catch (Exception e) {
        e.printStackTrace();
      }
    }

    /**
     * add Content
     */
    public void outFileAddContent(String content, String filePath) {
      File file = new File(filePath);
      file.getParentFile().mkdirs();
      String oldContent = u.readFile(filePath);
      try (PrintWriter writer = new PrintWriter(new BufferedWriter(new OutputStreamWriter(new FileOutputStream(filePath), StandardCharsets.UTF_8)))) {
        writer.println(oldContent + "\n" + content);
      } catch (Exception e) {
        e.printStackTrace();
      }
    }

    /**
     * @param content -
     * @param filePath -
     */
    public void outFileIFNotExist(String content, String filePath) {

      File file = new File(filePath);

      if (!file.exists()) {
        file.getParentFile().mkdirs();
      } else {
        System.out.println("Already exists:" + filePath);
        return;
      }

      try (PrintWriter writer = new PrintWriter(new BufferedWriter(new OutputStreamWriter(new FileOutputStream(filePath), StandardCharsets.UTF_8)))) {
        writer.println(content);
      } catch (Exception e) {
        e.printStackTrace();
      }

    }

    /**
     * 复制目录（含子目录/文件）
     *
     * @param source 源目录
     * @param target 目标目录
     * @param overwrite true: 覆盖已存在文件；false: 已存在则跳过
     * 
     *        <pre>
     *        Path src = Paths.get("/path/to/sourceDir");
     *        Path dst = Paths.get("/path/to/destDir");
     *        boolean overwrite = true; // true=覆盖, false=不覆盖
     *        copyDirectory(src, dst, overwrite);
     *        System.out.println("复制完成：" + dst);
     *        </pre>
     */
    public void copyDirectory(String from, String to, boolean overwrite) {
      Path source = Paths.get(from);
      Path target = Paths.get(to);
      try {
        // 创建目标根目录
        Files.createDirectories(target);

        File file = new File(from);
        // 存在しないの場合、戻る
        if (!file.exists()) {
          return;
        }
        // 如果不想跟随符号链接，可传入 EnumSet.of(FileVisitOption.FOLLOW_LINKS) 以外的配置
        Files.walkFileTree(source, EnumSet.noneOf(FileVisitOption.class), Integer.MAX_VALUE, new SimpleFileVisitor<Path>() {

          @Override
          public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) throws IOException {
            Path relative = source.relativize(dir);
            Path targetDir = target.resolve(relative);
            Files.createDirectories(targetDir);
            return FileVisitResult.CONTINUE;
          }

          @Override
          public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
            Path relative = source.relativize(file);
            Path targetFile = target.resolve(relative);

            // 复制选项：是否覆盖
            CopyOption[] options = overwrite ? new CopyOption[] {StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.COPY_ATTRIBUTES}
                : new CopyOption[] {StandardCopyOption.COPY_ATTRIBUTES};

            if (!overwrite && Files.exists(targetFile)) {
              // 已存在且不覆盖 → 跳过
              // 可按需打印日志：
              System.out.println("is exit file: " + targetFile);
              return FileVisitResult.CONTINUE;
            }
            // 确保父目录存在（理论上 preVisitDirectory 已创建，这里是双保险）
            Files.createDirectories(targetFile.getParent());
            Files.copy(file, targetFile, options);
            return FileVisitResult.CONTINUE;
          }

          @Override
          public FileVisitResult visitFileFailed(Path file, IOException exc) throws IOException {
            // 这里根据需求选择抛出或忽略
            throw exc;
          }
        });
      } catch (IOException e) {
        e.printStackTrace();
      }

    }

    /**
     * 获取指定目录下的所有子文件夹（以 File 形式返回）
     *
     * @param dirPath 目录路径，例如：C:\\opt\\API
     * @return 文件夹 File 列表
     * @throws IOException
     */
    public static List<File> listDir(String dirPath) throws IOException {
      List<File> result = new ArrayList<>();
      Path path = Paths.get(dirPath);

      if (Files.exists(path) && Files.isDirectory(path)) {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(path)) {
          for (Path p : stream) {
            if (Files.isDirectory(p)) {
              // Path → File
              result.add(p.toFile());
            }
          }
        }
      }

      return result;
    }


    /**
     * 正規表現ではなくリテラルとして扱う
     *
     * @param line 内容
     * @param regex 正規表現ではなく
     * @return boolean
     */
    public boolean isContentMatchQuote(String line, String regex) {
      Pattern p = Pattern.compile(Pattern.quote(regex));
      Matcher m = p.matcher(line);
      if (m.find()) {
        return true;
      }
      return false;
    }

    /**
     * 正規表現ではなくリテラルとして扱う
     *
     * @param line 内容
     * @param regex 正規表現ではなく
     * @return boolean
     */
    public boolean isContentMatch(String line, String regex) {
      Pattern p = Pattern.compile(regex);
      Matcher m = p.matcher(line);
      if (m.find()) {
        return true;
      }
      return false;
    }

    /**
     * 匹配最后的内容
     */
    public String contentMatchGroup1LastContent(String line, String regex) {

      String lastId = null;
      Pattern idPat = Pattern.compile(regex);
      Matcher idM = idPat.matcher(line);
      while (idM.find()) {
        lastId = idM.group(1); // 循环覆盖，结束即为最后一次
      }
      return lastId;
    }

    /**
     * 正規表現ではなくリテラルとして扱う
     *
     * @param line 内容
     * @param regex 正規表現ではなく
     * @return String
     */
    public String contentMatchGroup1(String line, String regex) {
      Pattern p = Pattern.compile(regex);
      Matcher m = p.matcher(line);
      if (m.find()) {
        return m.group(1);
      }
      return null;
    }

    /**
     * 正規表現ではなくリテラルとして扱う
     *
     * @param line 内容
     * @param regex 正規表現ではなく
     * @return String
     */
    public String contentMatchGroupNum(String line, String regex, Integer groupNum) {
      Pattern p = Pattern.compile(regex);
      Matcher m = p.matcher(line);
      if (m.find()) {
        return m.group(groupNum);
      }
      return null;
    }

    /**
     * 正規表現ではなくリテラルとして扱う
     *
     * @param line 内容
     * @param regex 正規表現ではなく
     * @return String
     */
    public String contentMatch(String line, String regex) {
      Pattern p = Pattern.compile(regex);
      Matcher m = p.matcher(line);
      if (m.find()) {
        return m.group();
      }
      return null;
    }

    public String replaceAll(String content, Map<String, String> map) {
      for (String key : map.keySet()) {
        content = content.replaceAll(key, map.get(key));
      }
      return content;
    }


    /**
     * reflectToFieldValue(fieldName, value, obj);
     */
    public void reflectToFieldValue(String fieldName, String value, Object obj) {
      try {
        Class<?> clazz = obj.getClass();

        // 获取所有声明的字段（包括私有字段）
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
          field.setAccessible(true); // 允许访问私有字段
          if (fieldName.equals(field.getName())) {
            // System.out.println("====" + field);
            field.set(obj, value);
          }

        }
      } catch (IllegalArgumentException e) {
        e.printStackTrace();
      } catch (IllegalAccessException e) {
        e.printStackTrace();
      }
    }

    /**
     * lower4
     *
     * @param input -
     * @return -
     */
    public String lower(String input) {
      return lower(input, 1);
    }

    /**
     * lower
     *
     * @param input -
     * @return -
     */
    public String lower(String input, int num) {
      // int num = 4;
      if (input.length() <= num) {
        return input;
      }
      if (input == null || input.isEmpty()) {
        return input;
      }
      return input.substring(0, num).toLowerCase() + input.substring(num);
    }

    /**
     * up
     *
     * @param input -
     * @return -
     */
    public String up(String input) {
      if (input == null || input.isEmpty()) {
        return input;
      }
      return input.substring(0, 1).toUpperCase() + input.substring(1);
    }

    /**
     * 创建StringBuilder模板
     *
     * @param content -
     * @return -
     */
    public String getStringBuilderTemplate(String content) {
      // content = content.replaceAll("\"", "\\\"");
      StringBuilder sb = new StringBuilder();
      sb.append("StringBuilder sb=new StringBuilder();\n");
      StringReader sr = new StringReader(content);
      BufferedReader br = new BufferedReader(sr);
      String b = "";
      try {
        while ((b = br.readLine()) != null) {
          b = b.replace("\\", "\\\\");
          b = b.replaceAll("\"", "\\\\\"");

          sb.append("sb.append(\"" + b + "\\n\");\n");
        }
      } catch (IOException e) {
        e.printStackTrace();
      }
      sb.append("return sb.toString();");
      return sb.toString();
    }

    /**
     * 创建StringBuilder模板
     *
     * @param content -
     * @return -
     */
    public String getStringBuilderTemplate1(String content, String n) {
      // content = content.replaceAll("\"", "\\\"");
      StringBuilder sb = new StringBuilder();
      sb.append("StringBuilder sb=new StringBuilder();\n");
      StringReader sr = new StringReader(content);
      BufferedReader br = new BufferedReader(sr);
      String b = "";
      try {
        while ((b = br.readLine()) != null) {

          if (isContentMatch(b, "package([^;]+);")) {
            System.out.println(b);
            b = "package \" + Sub.class.getPackage().getName()+  \"" + n + ";";
            sb.append("sb.append(\"" + b + "\\n\");\n");
          } else {
            b = b.replace("\\", "\\\\");
            b = b.replaceAll("\"", "\\\\\"");
            sb.append("sb.append(\"" + b + "\\n\");\n");

          }

        }
      } catch (IOException e) {
        e.printStackTrace();
      }
      sb.append("return sb.toString();");
      return sb.toString();
    }

  }

  /**
   * Application
   */
  class Application {

    static Sub sub;
    static Test my;
    static Oracle o = Sub.o;
    static Util u = Sub.u;
    static Excel e = Sub.e;

    /**
     * 内容をSubIn.txtに入れる
     */
    public void createKeyWord(String keyName) {
      String content = u.readFile(as.getSubInTxtPath());
      u.outFile(content, "C:\\xsd-fts\\src\\main\\java\\com\\fujitsu\\keys\\sub\\" + keyName);
    }

  }

  /**
   * ApplicationSub
   */
  public class ApplicationSub {

    static Test my;
    static Template t = Sub.t;

    public void outSubLog(String content) {
      System.out.println(content);
      u.outFileAddContent(content, as.getSubOutLogPath());
    }

    /**
     * getZopPath("bak"); C:\zop\bak
     */
    public String getZopPath(String subPath) {
      String path = "C:\\zop\\" + subPath + "\\";
      return path;
    }

    /**
     * C:\opt\workspace\SBMAB002\src\test\resources\
     */
    public String getSrcTestResourcesPath() {
      String dbTestDir = Sub.u.getProjectFindFirstDir("src/test", "resources");
      return dbTestDir;
    }

    /**
     * C:\opt\workspace\SBMAB002\src\main\resources\
     */
    public String getSrcMainResourcesPath() {
      String dbTestDir = Sub.u.getProjectDir("src/main", "") + "resources" + File.separator;
      return dbTestDir;
    }

    /**
     * C:\opt\workspace\SBMAB002\test\inputData\SBMA\SBMAB002\db\testCase\
     */
    public String getTestDbTestCasePath() {
      String dbTestDir = Sub.u.getProjectFindFirstDir("test/inputData", "db");
      File dbTestCaseDir = Sub.u.getFindFirstDir(dbTestDir, "testCase");
      String baseDir = dbTestCaseDir.getAbsolutePath() + File.separator;
      return baseDir;
    }

    /**
     * C:\opt\workspace\SBMAB002\test\inputData\SBMA\SBMAB002\db\testCase\111111\
     */
    public String getTestDbTestCaseTestNumPath(String testNum) {
      String to = getTestDbTestCasePath() + testNum + File.separator;
      return to;
    }

    /**
     * C:\opt\workspace\SBMAB002\test\inputData\SBMA\SBMAB002\db\testCase\111111\input\
     */
    public String getTestDbTestCaseTestNumInputPath(String testNum) {
      String to = getTestDbTestCaseTestNumPath(testNum) + "input" + File.separator;
      return to;
    }

    /**
     * C:\opt\workspace\SBMAB002\test\inputData\SBMA\SBMAB002\db\testCase\111111\output\
     */
    public String getTestDbTestCaseTestNumOutputPath(String testNum) {
      String to = getTestDbTestCaseTestNumPath(testNum) + "output" + File.separator;
      return to;
    }

    /**
     * 
     */
    public String getSubOutTxtPath() {
      String path = u.getCurrentFilePath();
      return path + "SubOut.txt";
    }

    /**
     * 
     */
    public String getSubOutHTMLPath() {
      String path = u.getCurrentFilePath();
      return path + "SubOut.html";
    }

    /**
    * 
    */
    public String getSubOutLogPath() {
      String path = u.getCurrentFilePath();
      return path + "SubOut.log";
    }

    /**
     * 
     */
    public String getSubOutJsonPath() {
      String path = u.getCurrentFilePath();
      return path + "SubOut.json";
    }

    /**
    * 
    */
    public String getSubInTxtPath() {
      String path = u.getCurrentFilePath();
      return path + "SubIn.txt";
    }

    /**
    * 
    */
    public String getSubInTxtContent() {
      String content = u.readFile(as.getSubInTxtPath());
      return content;
    }

    /**
    * 
    */
    public String getSubOutTxtContent() {
      String content = u.readFile(as.getSubOutTxtPath());
      return content;
    }

    /**
    * 
    */
    public String getSubTesterPath() {
      String path = u.getCurrentFilePath();
      return path + "SubTester.java";
    }

    /**
    * 
    */
    public String getSubTesterContent() {
      String content = u.readFile(as.getSubInTxtPath());
      return content;
    }

    /**
     * 获取电文路径 C:\opt\workspace\SBMAB002\test\inputData\SBMA\SBMAB002\telegram\data\\
     */
    public String getTelegramDataPath() {
      String path = u.getProjectFindFirstDir("test", "telegram");
      path = u.getFindFirstDir(path, "data").getAbsolutePath() + File.separator;
      return path;
    }


    /**
     * 获取DB 期望值目錄下指定表路徑
     */
    public String getJsonDBExpectPath(String caseNum, String tableName) {
      // 获取json
      String path = u.getProjectFindFirstDir("test", "db");
      path = u.getFindFirstDir(path, "testCase").getAbsolutePath() + File.separator;

      path = path + caseNum + "\\expect\\" + tableName + ".json";

      return path;
    }

    /**
     * 获取telegram 期望值目錄下指期望值路徑 input
     */
    public String getJsonTelegramInputPath(String caseNum) {
      // 获取json
      String path = u.getProjectFindFirstDir("test", "telegram");
      path = u.getFindFirstDir(path, "testCase").getAbsolutePath() + File.separator;

      path = path + caseNum + "\\input\\input.json";

      return path;
    }

    /**
     * 获取telegram 期望值目錄下指期望值路徑 expect
     */
    public String getJsonTelegramExpectPath(String caseNum) {
      // 获取json
      String path = u.getProjectFindFirstDir("test", "telegram");
      path = u.getFindFirstDir(path, "testCase").getAbsolutePath() + File.separator;

      path = path + caseNum + "\\expect\\expect.json";

      return path;
    }


    public void javaKeyPatterh(String pattern, String path) {
      // String pattern = "this.outputEndTraceLog";
      try {
        List<File> dirList = u.listDir(path);
        for (File dir : dirList) {
          System.out.println(dir.getAbsolutePath());
          getJavaPattern(pattern, dir.getAbsolutePath());
        }
      } catch (IOException e) {
        // TODO Auto-generated catch block
        e.printStackTrace();
      }

    }

    public void getJavaPattern(String patteString, String projectpath) {
      StringBuilder sb = new StringBuilder();
      String javaFile = getJavaFile(projectpath);
      if (javaFile == null) {
        return;
      }
      // System.out.println(javaFile);
      List<String> list = u.readFileToList(javaFile);
      for (int i = 0; i < list.size(); i++) {
        int num = i + 1;
        String s = list.get(i);
        boolean b = u.isContentMatchQuote(s, patteString);
        if (b) {

          sb.append("<a href='" + javaFile + "' target='_blank'>" + "\n");
          sb.append("Line：" + num + "\n");
          sb.append("</a>\n");

          // sb.append(javaFile + "\n");
          // sb.append("Line：" + num + "\n");
          // sb.append(s + "\n");
          System.out.println(sb.toString());
        }
      }
      u.outFileAddContent(sb.toString(), as.getSubOutTxtPath());
    }

    /**
     * 主要なじゃヴぁファイルを取得する
     */
    public String getJavaFile(String projectpath) {
      File path = u.getFindFirstDir(projectpath, "src");
      if (path == null) {
        return null;
      }
      File f = u.getFindFirstDir(path.getAbsolutePath(), "biz");
      List<File> fileList = u.findFile(f.getAbsolutePath());
      for (File file : fileList) {
        // System.out.println(file.getAbsolutePath());
        return file.getAbsolutePath();
      }
      return null;
    }

    /**
     * 创建mapper文件夹
     */
    public void createMapperDir() {
      String resourcesMapperDir = as.getSrcMainResourcesPath() + "mapper";
      File f = new File(resourcesMapperDir);
      f.mkdirs();
    }

    /**
     * → SubOut.txt
     */
    public void testClearContent() {
      String content = c.testClear();
      u.outFile(content, as.getSubOutTxtPath());
    }

    /**
     * <pre>
     * t1.SEC_CREENTR01_SWIFT_CD AS secCreentr01SwiftCd　→　SubIn.txt
     * 
     * sql仕様書获取 interYmd12Dt　→　SubOut.txt
     * 
     * to SubOut.log
     * 
     * </pre>
     */
    public void sqlAsAndBeanFieldCompare() {

      StringBuilder sb = new StringBuilder();

      // sql获取 t1.SEC_CREENTR01_SWIFT_CD AS secCreentr01SwiftCd,
      String intxt = as.getSubInTxtContent();
      List<String> inList = u.readContentToList(intxt);
      Map<String, String> inmap = new HashMap<String, String>();

      for (String s : inList) {
        String g = u.contentMatchGroup1(s, "AS (.*)?").trim().replaceAll(",", "");

        inmap.put(g, g);

      }

      // 文档获取 interYmd12Dt
      String outtxt = as.getSubOutTxtContent();
      List<String> outList = u.readContentToList(outtxt);
      Map<String, String> outmap = new HashMap<String, String>();
      for (String outs : outList) {
        outs = outs.trim();
        outmap.put(outs, outs);
      }


      System.out.println("from doc");
      for (String outs : outList) {
        outs = outs.trim();
        String s = inmap.get(outs);
        if (s == null) {
          String message = "mapper in not:" + outs;
          System.out.println(message);
          sb.append(message + "\n");
        }

      }

      System.out.println("from mapper");
      for (String ins : inList) {
        String g = u.contentMatchGroup1(ins, "AS (.*)?").trim().replaceAll(",", "");
        String s = outmap.get(g);
        if (s == null) {
          String message = "doc in not:" + g;
          System.out.println(message);
          sb.append(message + "\n");
        }

      }

      u.outFile(sb.toString(), as.getSubOutLogPath());

    }



    /**
     * 新しいテストを生成する 生成新的测试方法 输出到
     */
    public void createNewTest(String fromNum, String toNum) {
      as.testCopyTestCase(fromNum, toNum, "db");
      as.testCopyTestCase(fromNum, toNum, "telegram");
      as.testCopyTestCase(fromNum, toNum, "message");
      getTestNewMethod(fromNum, toNum);
    }

    /**
     * 新しいメソッドを生成する
     * 
     * 在 sub中生成新方法
     * 
     */
    public void getTestNewMethod(String fromNum, String toNum) {

      StringBuilder newContent = new StringBuilder();
      StringBuilder sb = new StringBuilder();
      LinkedHashMap<String, String> map = getTesterContentSplit();

      int i = 0;
      for (String key : map.keySet()) {

        String content = map.get(key);
        if (u.isContentMatch(content, "test" + fromNum + "\\(\\)")) {
          newContent.append(content.replaceAll(fromNum, toNum) + "\n");
        }

        if (i == map.size() - 1) {
          sb.append(newContent.toString());
        }

        // sb.append(content + "\n");

        i = i + 1;
      }

      System.out.println(sb.toString());
      String s = u.readFile(as.getSubOutTxtPath());
      u.outFile(s + "\n" + sb.toString(), as.getSubOutTxtPath());
    }

    /**
     * javaのテストメソッドの内容を取得する分割
     */
    public LinkedHashMap<String, String> getTesterContentSplit() {

      String testJavaName = u.getProjectName() + "Tester.java";

      String filePath = u.getProjectFindFirstDir("src", "test");

      File f = u.findFile(filePath, testJavaName);

      String testJavaPath = f.getAbsolutePath();
      List<String> testList = u.readFileToList(testJavaPath);

      LinkedHashMap<String, String> linkedMap = new LinkedHashMap<String, String>();
      StringBuilder sb = new StringBuilder();

      for (int i = 0; i < testList.size(); i++) {

        String s = testList.get(i);

        if (s.isEmpty()) {
          sb.append("\n");
          linkedMap.put(String.valueOf(i), sb.toString());
          sb = new StringBuilder();
        } else {
          sb.append(s + "\n");
        }

      }

      if (sb.length() > 0) {
        // Key 用什么看你的需求：
        // 1) 用最后一行的下标：testList.size() - 1
        // 2) 或者用一个特殊 key：比如 "EOF"
        linkedMap.put(String.valueOf(testList.size() - 1), sb.toString());
        // linkedMap.put("EOF", sb.toString());
      }

      return linkedMap;

    }

    /**
     * 获取表定义并打印到SubOut.json
     * 
     */
    public void getTableDefinePrint(String tableName) {
      String dbDefinePath = u.getProjectFindFirstDir("test", "db");
      dbDefinePath = u.getFindFirstDir(dbDefinePath, "define").getAbsolutePath() + File.separator;
      String json = u.readFile(dbDefinePath + tableName + ".json");
      u.outFile(json, this.getSubOutJsonPath());
      System.out.println(json);
    }

    /**
     * 自分でinputのjsonデータを設定する
     * 
     * TB_FITACCOUNT "DEALER_ACC_CD": "11","NEWCO_CD": null,"NEW_KBN_ACC_CD": null,
     */
    public void setJsonDbInputValueChange(String testNum, String param) {

      String tableName = param.split("\s+")[0].trim();
      String fieldContent = param.replaceAll(tableName + "\s+", "");
      // "DEALER_ACC_CD": "11","NEWCO_CD": null,"NEW_KBN_ACC_CD": null,
      String[] fieldsContent = fieldContent.split(",");
      LinkedHashMap<String, String> linkedMap = new LinkedHashMap<String, String>();
      for (String s : fieldsContent) {
        // 逗号可选
        String regex = "\"([^\"]+)\"\\s*:\\s*(?:\"([^\"]*)\"|(?i)null)\\s*,?";
        String fieldName = u.contentMatchGroup1(s, regex);
        String value = u.contentMatchGroupNum(s, regex, 2);
        // System.out.println(fieldName + "====" + value);
        linkedMap.put(fieldName, value);
      }

      String path = u.getProjectFindFirstDir("test", "db");
      path = u.getFindFirstDir(path, "testCase").getAbsolutePath() + File.separator + testNum + File.separator + "input" + File.separator;
      String filePath = path + tableName + ".json";


      String json = u.readFile(filePath);

      @SuppressWarnings("unchecked")
      LinkedHashMap<String, List<?>> payload = JsonUtils.fromJson(json, LinkedHashMap.class);

      @SuppressWarnings("unchecked")
      List<LinkedHashMap<String, String>> rows = (List<LinkedHashMap<String, String>>) payload.get("tableData");
      for (LinkedHashMap<String, String> map : rows) {
        for (String key : linkedMap.keySet()) {
          if (map.get(key) != null) {
            map.put(key, linkedMap.get(key));
          }
        }
      }
      u.outFile(JsonUtils.toJson(payload), filePath);
    }

    /**
     * 
     */
    public void clearLog() {
      String filePath = u.getProjectDir("log") + "app.log";
      u.outFile("", filePath);
    }

    /**
     * testCopyTestCase("010100", "010147", "db");
     * 
     * testCopyTestCase("010136", "010147", "telegram");
     */
    public void testCopyTestCase(String fromNum, String toNum, String testDir) {
      String dbTestDir = u.getProjectFindFirstDir("test/inputData", testDir);
      File dbTestCaseDir = u.getFindFirstDir(dbTestDir, "testCase");

      String baseDir = dbTestCaseDir.getAbsolutePath() + File.separator;
      String from = baseDir + fromNum;
      String to = baseDir + toNum;
      u.copyDirectory(from, to, false);
    }

    /**
     * 初始化 SubOut SubIn SubTester
     */
    public void testInit(String[] fileNames) {
      String path = u.getCurrentFilePath();

      u.outFileIFNotExist("", path + "SubOut.txt");
      u.outFileIFNotExist("", path + "SubIn.txt");
      u.outFileIFNotExist(t.getSubTester(), path + "SubTester.java");

      for (String fileName : fileNames) {
        try {

          String javaContent = (String) t.getClass().getMethod("get" + fileName).invoke(t);
          u.outFileIFNotExist(javaContent, path + "sub/" + fileName + ".java");
        } catch (IllegalAccessException e) {
          e.printStackTrace();
        } catch (InvocationTargetException e) {
          e.printStackTrace();
        } catch (NoSuchMethodException e) {
          e.printStackTrace();
        } catch (SecurityException e) {
          e.printStackTrace();
        }

      }

      // u.outFileIFNotExist(t.getSubTester(), path + "SubTester.java");
      // u.outFileIFNotExist(t.getJsonUtils(), path + "sub/JsonUtils.java");
      // u.outFileIFNotExist(t.getKeyUI(), path + "sub/KeyUI.java");

    }


    /**
     * 自分で定義したバイク
     */
    public void subBak(String content, String backName) {
      String formatted = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"));
      String bakPath = getZopPath("bak") + backName + "_" + formatted;
      u.outFile(content, bakPath);
    }

    /**
    *
    */
    public void subBak(String[] fileNames) {

      as.testInit(fileNames);

      StringBuilder sb = new StringBuilder();
      sb.append(subContent());

      sb.append("  /**\n");
      sb.append("   * テンプレートクラス\n");
      sb.append("   */\n");
      sb.append("  class Template {\n");


      // 创建方法
      sb.append("    /**\n");
      sb.append("     * \n");
      sb.append("     */\n");
      sb.append("    public String getSubTester() {\n");
      String path = u.getCurrentFilePath() + "SubTester.java";
      String content = u.readFile(path);
      String s = u.getStringBuilderTemplate1(content, "");
      sb.append(s + "\n");
      sb.append("    }\n");


      // 遍历模板
      for (String fileName : fileNames) {

        sb.append("    /**\n");
        sb.append("     * \n");
        sb.append("     */\n");
        sb.append("    public String get" + fileName + "() {\n");

        path = u.getCurrentFilePath() + "sub/" + fileName + ".java";
        content = u.readFile(path);

        s = u.getStringBuilderTemplate1(content, ".sub");
        sb.append(s + "\n");

        sb.append("    }\n");

      }

      sb.append("}\n");
      sb.append("}\n");

      String bakPath = getZopPath("bak") + "Sub.java";
      u.outFile(sb.toString(), bakPath);

    }

    /**
     * 获取除了模板类之外的数据
     */
    public String subContent() {

      StringBuilder sb = new StringBuilder();

      String subJavaPath = u.getCurrentFilePath() + "Sub.java";
      List<String> list = u.readFileToList(subJavaPath);

      // 获取最后一个【テンプレートクラス】所在行
      int maxLine = 0;
      int i = 0;
      for (String s : list) {
        if (u.isContentMatch(s, "テンプレートクラス")) {
          maxLine = i;
        }
        i = i + 1;
      }

      // 获取现有模板
      // for (int j = maxLine - 1; j < list.size(); j++) {
      // // String s = list.get(j);
      // }

      // 现有模板之外
      for (int j = 0; j < maxLine - 1; j++) {
        String s = list.get(j);
        sb.append(s + "\n");
      }

      return sb.toString();
    }

    /**
     * SubOut.txtに格納する
     * 
     */
    public void subTesterToTemplate() {
      String content = getSubTesterContent();
      String s = u.getStringBuilderTemplate(content);
      u.outFile(s, this.getSubOutTxtPath());
    }

    /**
     * <pre>
     選択したテンプレートをすべて「DBのtestcaseのnum」フォルダに格納する
     getJsonTemplateToAllTestCaseDir("TB_FITTERMNOTICEINFO","010");
     * </pre>
     */
    public void getJsonTemplateToAllDbTestCaseDir(String templateName, String casePre) {
      // 放入模板
      String content = u.readFile(getZopPath("tableJson") + templateName + ".json");
      for (int i = 100; i < 147; i++) {
        String caseNum = casePre + i;
        String path = u.getProjectFindFirstDir("test", "db");
        path = u.getFindFirstDir(path, "testCase").getAbsolutePath() + File.separator;
        path = path + caseNum + "\\expect\\TB_FITTERMNOTICEINFO.json";
        System.out.println(path);
        u.outFileIFNotExist(content, path);
      }
    }

    /**
     * <pre>
     * 获取真实件数并放入到期望json中
     * ログの中の “but was” キーワードを基に、JSON の正しい件数を取得する
     * </pre>
     */
    public void setJsonDbExpectRealLength() {

      List<String> saBunnContents = getLogSaBunnALi();

      for (String content : saBunnContents) {

        String caseNum = u.contentMatchGroup1(content, "テストケース\\[([^]]+)\\]【テーブル比較結果】");
        String jsonWas = u.contentMatchGroup1(content, "but was:<(\\d+)>");
        String tableName = u.contentMatchGroup1(content, "table=([^\\)]+)");
        if (jsonWas != null) {

          String path = getJsonDBExpectPath(caseNum, tableName);
          String json = u.readFile(path);

          @SuppressWarnings("unchecked")
          LinkedHashMap<String, List<?>> payload = JsonUtils.fromJson(json, LinkedHashMap.class);

          @SuppressWarnings("unchecked")
          List<LinkedHashMap<String, String>> rows = (List<LinkedHashMap<String, String>>) payload.get("tableData");

          List<LinkedHashMap<String, String>> rowsNew = new ArrayList<LinkedHashMap<String, String>>();

          if (rows.size() == 0) {

            String dbJsonContent = setDbJsonDefineToMap(tableName);
            u.outFile(dbJsonContent, path);

          } else {
            for (int i = 0; i < Integer.valueOf(jsonWas); i++) {
              rowsNew.add(rows.get(0));
            }
            payload.put("tableData", rowsNew);
            u.outFile(JsonUtils.toJson(payload), path);
          }

        }
      }
    }

    /**
     * 初始化数据Db json数据
     */
    public String setDbJsonDefineToMap(String tableName) {
      // String tableName = "TB_FITTERMNOTICEINFO";

      String dbDefinePath = u.getProjectFindFirstDir("test", "db");
      dbDefinePath = u.getFindFirstDir(dbDefinePath, "define").getAbsolutePath() + File.separator;

      String json = u.readFile(dbDefinePath + tableName + ".json");

      @SuppressWarnings("unchecked")
      LinkedHashMap<String, List<?>> payload = JsonUtils.fromJson(json, LinkedHashMap.class);

      @SuppressWarnings("unchecked")
      List<LinkedHashMap<String, String>> rows = (List<LinkedHashMap<String, String>>) payload.get("tableStructure");

      LinkedHashMap<String, Object> newPayload = new LinkedHashMap<String, Object>();
      newPayload.put("physicalTableName", tableName);

      List<LinkedHashMap<String, String>> tableDataList = new ArrayList<>();
      LinkedHashMap<String, String> newLinkedMap = new LinkedHashMap<String, String>();

      tableDataList.add(newLinkedMap);
      newPayload.put("tableData", tableDataList);

      for (LinkedHashMap<String, String> obj : rows) {

        String physicalName = obj.get("physicalName");
        newLinkedMap.put(physicalName, null);
      }
      return JsonUtils.toJson(newPayload);
    }

    /**
     * <pre>
      获取日志中数据不一致的值,获取实际值进行修正并放入到期望json中
      シャードログで不一致が発生した場合、expect には実際の値を入れる。
     * </pre>
     */
    public void setJsonDbExpectRowI() {

      List<String> saBunnContents = getLogSaBunnALi();
      for (String content : saBunnContents) {

        LinkedHashMap<String, List<String>> rowMap = getLogDbRowI(content);

        String caseNum = rowMap.get("testNum").get(0);
        String tableName = rowMap.get("tableName").get(0);

        if (tableName.isEmpty()) {
          System.out.println("not row[i] :" + content);
          continue;
        }

        String path = getJsonDBExpectPath(caseNum, tableName);
        String json = u.readFile(path);

        @SuppressWarnings("unchecked")
        LinkedHashMap<String, List<?>> payload = JsonUtils.fromJson(json, LinkedHashMap.class);

        @SuppressWarnings("unchecked")
        List<LinkedHashMap<String, String>> rows = (List<LinkedHashMap<String, String>>) payload.get("tableData");

        rowMap.remove("testNum");
        rowMap.remove("tableName");
        for (String key : rowMap.keySet()) {
          List<String> rowNumContentList = rowMap.get(key);
          for (String s : rowNumContentList) {
            String regex = "結果項目\\[([^.]+)\\.([^]]+)\\]=\\[([^]]+)\\]";
            // String tableName = u.contentMatchGroup1(s, regex);
            String fieldName = u.contentMatchGroupNum(s, regex, 2);
            String value = u.contentMatchGroupNum(s, regex, 3);
            if (value != null) {
              if (value.length() < 30) {
                // 2025-12-01 → 2025/12/01
                value = value.replaceAll("(\\d{4})-(\\d{2})-(\\d{2})", "$1/$2/$3");
                // 2026/01/13 00:00:00.0 → 2026/01/13 00:00:00
                value = value.replaceAll("((\\d{4})/(\\d{2})/(\\d{2})\\s+\\d{2}:\\d{2}:\\d{2})\\.0+\\b", "$1");
              }
            }
            // Row[1], 結果項目[TB_FITISSUERNXT.NEWCO_CD]=[null],
            // 期待項目[TB_FITISSUERNXT.NEWCO_CD]=[null ],
            // 差異内容[項目値不一致]
            // 解决null字段
            if (value.trim().equals("null")) {
              value = null;
            }
            // System.out.println(caseNum + "=" + tableName + "=" + fieldName + "=" +
            // value);

            // 第i组数据
            LinkedHashMap<String, String> obj = rows.get(Integer.valueOf(key));
            obj.put(fieldName, value);
          }
        }
        u.outFile(JsonUtils.toJson(payload), path);
      }
    }

    /**
     * <pre>
    content:
     2026-03-02 12:59:00.540 [main] ERROR jp.co.skywave.ut.db.util.UTDbUtil - テストケース[010107]テーブル比較結果=【★★★差分あり★★★】
    - 【テーブル期待差分：項目差分一覧】
    - Row[0], 結果項目[TB_FITTERMNOTICEINFO.BIZ_FUNC_ID_CD]=[ITMAB203], 期待項目[TB_FITTERMNOTICEINFO.BIZ_FUNC_ID_CD]=[999     ], 差異内容[項目値不一致]
    - Row[0], 結果項目[TB_FITTERMNOTICEINFO.MGMT_TER_NOTICE_TYP]=[E], 期待項目[TB_FITTERMNOTICEINFO.MGMT_TER_NOTICE_TYP]=[999], 差異内容[項目値不一致]
    - Row[0], 結果項目[TB_FITTERMNOTICEINFO.USR_ID_CD]=[user], 期待項目[TB_FITTERMNOTICEINFO.USR_ID_CD]=[999], 差異内容[項目値不一致]
    - Row[1], 結果項目[TB_FITTERMNOTICEINFO.BIZ_FUNC_ID_CD]=[ITMAB203], 期待項目[TB_FITTERMNOTICEINFO.BIZ_FUNC_ID_CD]=[999     ], 差異内容[項目値不一致]
    - Row[1], 結果項目[TB_FITTERMNOTICEINFO.MGMT_TER_NOTICE_TYP]=[E], 期待項目[TB_FITTERMNOTICEINFO.MGMT_TER_NOTICE_TYP]=[999], 差異内容[項目値不一致]
    - Row[1], 結果項目[TB_FITTERMNOTICEINFO.USR_ID_CD]=[user], 期待項目[TB_FITTERMNOTICEINFO.USR_ID_CD]=[999], 差異内容[項目値不一致]
     * 
     * 按照  Row[0] Row[1] 放入map
     * 
     * </pre>
     */
    private LinkedHashMap<String, List<String>> getLogDbRowI(String content) {
      // テストケース[010138]テーブル
      String testNum = u.contentMatchGroup1(content, "テストケース\\[([^]]+)\\]テーブル");

      List<String> list = u.readContentToList(content);

      String tableName = "";
      LinkedHashMap<String, List<String>> mapRows = new LinkedHashMap<String, List<String>>();

      for (String s : list) {

        if (s.contains("Row[")) {
          // - Row[0] → Row[0]
          String row = s.split(",")[0].trim().split("-")[1].trim();
          String regex = "結果項目\\[([^.]+)\\.([^]]+)\\]=\\[([^]]+)\\]";
          tableName = u.contentMatchGroup1(s, regex);
          // Row[0] → 0
          // Row[12] → 12
          String rowNum = u.contentMatchGroup1(row, "Row\\[([^]]+)\\]");

          List<String> rowNumList = mapRows.get(rowNum);
          if (rowNumList == null) {
            rowNumList = new ArrayList<String>();
          }
          rowNumList.add(s);

          mapRows.put(rowNum, rowNumList);
        }
      }

      List<String> testNumList = new ArrayList<String>();
      List<String> tableNameList = new ArrayList<String>();

      testNumList.add(testNum);
      tableNameList.add(tableName);

      mapRows.put("testNum", testNumList);
      mapRows.put("tableName", tableNameList);

      return mapRows;
    }

    /**
     * <pre>
     log差分分片
     差分あり
     获取日志差分分片，每一个片段都用list来分片
     * </pre>
     */
    public List<String> getLogSaBunnALi() {

      String path = u.getProjectDir("log") + "app.log";
      List<String> list = u.readFileToList(path);
      boolean start = false;

      StringBuilder sb = new StringBuilder();
      List<String> blocks = new ArrayList<String>();

      for (String line : list) {

        if (u.isContentMatch(line, "差分あり")) {
          start = true;
        }

        if (start) {
          sb.append(line + "\n");

          // 空白行结束片段（含仅空格/制表/全角空白等）
          if (line.trim().isEmpty()) {
            // 避免只把一个空行当成片段
            if (sb.length() > 1) { // 至少有一个 '\n'
              blocks.add(sb.toString());
            }
            sb.setLength(0);
            start = false;
          }

        }

      }
      return blocks;
    }

    /**
     * 检测电文是否有差分 true 有 需要变更
     */
    public Boolean isTelegramValueChange() {
      List<String> list = as.getLogSaBunnALi();
      for (String s : list) {
        // 获取实际值
        LinkedHashMap<String, String> colMap = getLogTelegramCol(s);
        // 测试号 没有有差分的电文
        String testNum = colMap.get("testNum");
        if (testNum == null) {
          // 非电文继续
          return false;
        }
      }
      return true;
    }

    /**
     * 電文データ修正
     */
    public void setJsonTelegramValueChange() {
      List<String> list = as.getLogSaBunnALi();
      for (String s : list) {
        // 获取实际值
        LinkedHashMap<String, String> colMap = getLogTelegramCol(s);
        // 测试号 没有有差分的电文
        String testNum = colMap.get("testNum");
        if (testNum == null) {
          // 非电文继续
          continue;
        }

        // C:\opt\PT\SBMA\SBMAA205\test\inputData\SBMA\SBMAA205\telegram\data\SBTAI205_1001.json
        // 获取电文路径
        String getTelegramJsonPath = getTelegramDataColPath(s, colMap);
        String json = u.readFile(getTelegramJsonPath);

        colMap.remove("testNum");
        colMap.remove("telegramNum");
        for (String key : colMap.keySet()) {
          // System.out.println(key + ": " + colMap.get(key));
          String regex = "(\"" + key + "\"\\s*:\\s*)\"([^\"]*)\"";
          json = json.replaceFirst(regex, "$1\"" + colMap.get(key) + "\"");
        }
        u.outFile(json, getTelegramJsonPath);
      }
    }

    /**
     * 初始化创建电文期望值文件
     */
    public void createInitTelegramExpectInputSet(String caseNum) {
      String expectFilePath = as.getJsonTelegramExpectPath(caseNum);
      File f = new File(expectFilePath);
      String json = c.telegramExpectJson();
      u.outFileIFNotExist(json, f.getAbsolutePath());

      String inputFilePath = as.getJsonTelegramInputPath(caseNum);
      f = new File(inputFilePath);
      json = c.telegramInputJson();
      u.outFileIFNotExist(json, f.getAbsolutePath());
    }

    /**
     * 清空电文期望值
     */
    public void createInitTelegramExpectInputNull(String caseNum) {
      String expectFilePath = as.getJsonTelegramExpectPath(caseNum);
      File f = new File(expectFilePath);
      String json = c.telegramExpectJson();
      u.outFile(json, f.getAbsolutePath());

      String inputFilePath = as.getJsonTelegramInputPath(caseNum);
      f = new File(inputFilePath);
      json = c.telegramInputJson();
      u.outFile(json, f.getAbsolutePath());
    }



    /**
     * <pre>
     * 找出匹配电文
     * 操作方法，不断的执行test 然后生成日志，运行此方法如果不是匹配电文则枚举下一个
     * 然后再执行test 依次类推知道电文是匹配的电文
     * 初始化电文,条件是电文文件为空,这样会显示没有匹配电文的数量
     * 根据数量穷举需要的电文
      {
        "expectTestDataIDs": []
      }
      C:\opt\workspace\SBMAB002\test\inputData\SBMA\SBMAB002\telegram\define 放入模板数据 
      {"expectTestDataIDs":["0_SBJMO011.json","0_SBJMO011.json","0_SBJMO011.json","0_SBJMO011.json","0_SBJMO011.json","0_SBJMO011.json"]}
      使得在匹配时能匹配的到，此程序着重寻找能够对应上的电文只要对应上了就成功了
     * </pre>
     */
    public void findJsonTelegramOK() {

      // 获取没有配置期望值文件的电文数量 ← log
      LinkedHashMap<String, String> telegramMap = new LinkedHashMap<String, String>();
      List<String> list = as.getLogSaBunnALi();
      for (String content : list) {
        // テストケース[000001]電文インデックス[0]
        String regex = "テストケース\\[([^]]+)\\]電文インデックス\\[([^]]+)\\]";
        String caseNum = u.contentMatchGroup1(content, regex);
        // 电文号对应测试的电文顺序0号 1号
        String index = u.contentMatchGroupNum(content, regex, 2);

        // 如果找到了匹配,说明还没有电文文件数据
        regex = "【差異】NotDefine|【差異】ChannelDiff|NG:Length";
        Boolean haveTel = u.isContentMatch(content, regex);
        // System.out.println(haveTel + "==" + index + "==" + testNum);
        if (haveTel == true && index != null) { // 等于空的不属于电文
          telegramMap.put(index, caseNum);
        }
      }

      // 获取当前电文模板列表
      String telegramTemplatePath = "C:\\zop\\telegram\\" + u.getProjectName();
      List<File> fileList = u.findFile(telegramTemplatePath);
      if (fileList.size() == 0) {
        String msg = "没有配置电文模板：" + telegramTemplatePath;
        u.outFile(msg, as.getSubOutLogPath());
      }

      // 获取测试号
      String caseNum = null;
      for (String index : telegramMap.keySet()) {
        caseNum = telegramMap.get(index);
      }

      if (caseNum == null) {
        return;
      }

      // 获取期望值文件 进行初始化操作
      String expectFilePath = as.getJsonTelegramExpectPath(caseNum);
      File f = new File(expectFilePath);
      String json = null;
      if (f.exists()) {
        json = u.readFile(expectFilePath);
      } else {
        json = c.telegramExpectJson();
      }

      @SuppressWarnings("unchecked")
      LinkedHashMap<String, List<String>> payload = JsonUtils.fromJson(json, LinkedHashMap.class);
      List<String> cols = payload.get("expectTestDataIDs");

      for (String index : telegramMap.keySet()) {
        // System.out.println(index + "==" + telegramMap.get(index));
        if (Integer.valueOf(index) < cols.size()) {
          // 如果没有这个索引则增加这个索引
          String telegramFileName = cols.get(Integer.valueOf(index));

          for (int i = 0; i < fileList.size(); i++) {
            String fileName = fileList.get(i).getName().split("\\.")[0];
            String fileNextName = fileName;
            // 换下一个电文 上一个电文不是
            if ((i + 1) != fileList.size()) {
              fileNextName = fileList.get(i + 1).getName().split("\\.")[0];
            }
            if (fileName.equals(telegramFileName)) {
              cols.set(Integer.valueOf(index), fileNextName);
              break;
            }
          }
        } else {
          // 如果没有则初始化
          cols.add(0, fileList.get(0).getName().split("\\.")[0]);
        }
      }
      // System.out.println(JsonUtils.toJson(payload));
      u.outFileAddContent(JsonUtils.toJson(payload), as.getSubOutLogPath());
      u.outFile(JsonUtils.toJson(payload), expectFilePath);


      // String caseNum = "";

      // 电文数据初始化

      // 获取电文列表

    }

    /**
     * 生成匹配电文数据 放入到 C:\opt\workspace\SBMAB002\test\inputData\SBMA\SBMAB002\telegram\data 中
     */
    public void createJsonTelegramOK(String caseNum) {

      String expectFilePath = as.getJsonTelegramExpectPath(caseNum);
      String json = u.readFile(expectFilePath);

      @SuppressWarnings("unchecked")
      LinkedHashMap<String, List<String>> payload = JsonUtils.fromJson(json, LinkedHashMap.class);
      List<String> cols = payload.get("expectTestDataIDs");

      String path = Sub.as.getTelegramDataPath();
      for (int i = 0; i < cols.size(); i++) {
        String telegramName = cols.get(i);
        String telegramFilePath = path + telegramName + ".json";
        String telegramJson = u.readFile(telegramFilePath);
        String regex = "(\"testDataID\"\\s*:\\s*)\"([^\"]*)\"";

        String newTelegramName = caseNum + "_" + i + "_" + telegramName.split("_")[1];
        telegramJson = telegramJson.replaceFirst(regex, "$1\"" + newTelegramName + "\"");
        u.outFileIFNotExist(telegramJson, path + newTelegramName + ".json");
        cols.set(i, newTelegramName);
      }
      u.outFile(JsonUtils.toJson(payload), expectFilePath);
    }

    /**
     * 获取真实引用电文jsoPath
     */
    public String getTelegramDataColPath(String s, LinkedHashMap<String, String> colMap) {

      // 测试号
      String testNum = colMap.get("testNum");
      // 电文号对应测试的电文顺序0号 1号
      String telegramNum = colMap.get("telegramNum");

      String path = as.getJsonTelegramExpectPath(testNum);

      // {
      // "expectTestDataIDs": [
      // "SBTAO205_1001"
      // ]
      // }
      String json = u.readFile(path);
      List<String> cols = getTelegramExpectList(json);

      // SBTAO205_1001 获取电文号
      String col = cols.get(Integer.valueOf(telegramNum)).replaceAll("\\\"", "");

      path = u.getProjectFindFirstDir("test", "telegram");
      path = u.getFindFirstDir(path, "data").getAbsolutePath() + File.separator;
      path = path + col + ".json";
      return path;
    }


    /**
     * 获取期待值中电文索引列表
     */
    public List<String> getTelegramExpectList(String json) {
      @SuppressWarnings("unchecked")
      LinkedHashMap<String, List<String>> payload = JsonUtils.fromJson(json, LinkedHashMap.class);

      List<String> cols = payload.get("expectTestDataIDs");

      return cols;
    }

    /**
     * <pre>
    content:
    2026-03-03 18:05:36.375 [main] ERROR j.c.s.u.t.c.UTTelegramComparator - テストケース[000001]電文インデックス[0] 比較結果=[★★★差分あり★★★]
    - 【電文期待差分：サマリー】
    - 【電文ID】SBTAO205
    - 【差異】Item
    - 【出力チャネル】期待値=[(なし)], 実値=[(なし)]
    - 【電文期待差分：詳細】
    - Col[0]:比較結果=[NG:Value]項目ID=[resultCode], 項目名=[処理結果], 項目パス=[restResHeader-resultCode], 期待値=[1], 実値=[9], 期待値バイト長=[1], 実値バイト長=[1]
    - Col[1]:比較結果=[NG:Value]項目ID=[messageId], 項目名=[メッセージID], 項目パス=[restResHeader-messageId], 期待値=[null], 実値=[SBA205114E], 期待値バイト長=[0], 実値バイト長=[10]
    - Col[2]:比較結果=[NG:Value]項目ID=[message], 項目名=[メッセージ文言], 項目パス=[restResHeader-message], 期待値=[null], 実値=[エラーレコードあり], 期待値バイト長=[0], 実値バイト長=[27]
    - Col[3]:比較結果=[OK:NotTarget]項目ID=[allCount], 項目名=[総投入件数], 項目パス=[allCount], 期待値=[1], 実値=[1], 期待値バイト長=[1], 実値バイト長=[1]
    - Col[4]:比較結果=[OK:NotTarget]項目ID=[normEndNum], 項目名=[正常終了件数], 項目パス=[normEndNum], 期待値=[1], 実値=[0], 期待値バイト長=[1], 実値バイト長=[1]
    - Col[5]:比較結果=[OK:NotTarget]項目ID=[errNum], 項目名=[エラー件数], 項目パス=[errNum], 期待値=[0], 実値=[1], 期待値バイト長=[1], 実値バイト長=[1]
    - Col[6]:比較結果=[NG:Length]項目ID=[chkErrInfo], 項目名=[チェックエラー情報], 項目パス=[chkErrInfo], 期待値=[配列長:0], 実値=[配列長:1], 期待値バイト長=[0], 実値バイト長=[0]
     * 
     * 单个电文
     * 
     * </pre>
     */
    private LinkedHashMap<String, String> getLogTelegramCol(String content) {

      // テストケース[000001]電文インデックス[0]
      String regex = "テストケース\\[([^]]+)\\]電文インデックス\\[([^]]+)\\]";

      String testNum = u.contentMatchGroup1(content, regex);

      // 电文号对应测试的电文顺序0号 1号
      String telegramNum = u.contentMatchGroupNum(content, regex, 2);

      List<String> list = u.readContentToList(content);

      LinkedHashMap<String, String> mapCols = new LinkedHashMap<String, String>();

      for (String s : list) {

        if (s.contains("Col[")) {

          // - Row[0] → Row[0]
          String row = s;
          regex = "項目名=\\[([^]]+)\\]";
          String colName = u.contentMatchGroup1(row, regex);
          regex = "実値=\\[([^]]+)\\]";
          String value = u.contentMatchGroup1(row, regex);

          mapCols.put(colName, value);
          // System.out.println(colName + "==" + value);

        }
      }

      mapCols.put("testNum", testNum);
      mapCols.put("telegramNum", telegramNum);

      return mapCols;
    }

    /**
     * エラーメッセージを設定する
     */
    public void setTestMessage() {
      List<String> list = as.getLogSaBunnALi();
      String dbTestDir = u.getProjectFindFirstDir("test/inputData", "message");
      File dbTestCaseDir = u.getFindFirstDir(dbTestDir, "testCase");
      String messageDir = dbTestCaseDir.getAbsolutePath() + File.separator;

      for (String s : list) {
        String caseNum = u.contentMatchGroup1LastContent(s, "テストケース\\[([^]]+)\\]メッセージ種別");
        String code = u.contentMatchGroup1LastContent(s, "実値Code=\\[([^]]+)\\],");
        String id = u.contentMatchGroup1LastContent(s, "実値ID=\\[([^]]+)\\],");
        String value = u.contentMatchGroup1LastContent(s, "実値文言=\\[(.*)\\]");

        String messagePath = messageDir + caseNum + File.separator;

        String message = "[2026-03-01 13:33:33] [Error] [TCID: " + caseNum + "] [MSGCD: " + code + "] [CID: " + id + "] \"" + value + "\"";
        String expectApllogFilePath = messagePath + "expect\\expect_apllog.log";
        String expectSyslogFilePath = messagePath + "expect\\expect_syslog.log";

        u.outFile(message, expectApllogFilePath);
        u.outFile(message, expectSyslogFilePath);

      }
    }

    /**
     * 从一条sql语句设置一个case
     */
    public void setOneCaseFromSql(String caseNum, Class<?> testClass) {
      // SubSql.setOneCaseFromSql(caseNum, testClass);
      try {
        // 1. 获取 SubSql 类
        Class<?> clazz = Class.forName(Sub.class.getPackage().getName() + ".sub.SubSql");
        // 2. 获取方法（参数类型必须完全一致）
        Method method = clazz.getMethod("setOneCaseFromSql", String.class, Class.class);
        // 3. 调用静态方法，invoke 的第一个参数传 null
        method.invoke(null, caseNum, testClass);
      } catch (Exception e) {
        throw new RuntimeException(e);
      }

    }

  }



  /**
   * Oracle
   */
  class Oracle {

    static Sub sub;
    static B b;
    static Excel e;
    static Test my;

    private String dbUrl = "jdbc:postgresql://localhost:5432/jasdec_sbg_f";
    private String userName = "postgres";
    private String pass = "jasdec2026";
    // private String dbName = "jasdec_sbg_f";

    /**
     * 
     */
    public void test() {
      String sql = "select * from TB_FSBREQUEST";
      List<B> list = queryAsObjects(sql);
      for (B b : list) {
        System.out.println(b.columnName);
      }

    }

    // create insert update
    @SuppressWarnings("unused")
    public String execute(String sql) throws SQLException {
      Connection conn = null;
      PreparedStatement pstmt = null;
      conn = getConnection(); // コネクションを取得
      conn.setAutoCommit(false); // 自動提出禁止}

      Statement stmt = conn.createStatement();
      stmt.execute(sql);
      stmt.close();
      conn.commit();
      return "true";
    }

    private Connection getConnection() throws SQLException {
      return DriverManager.getConnection(dbUrl, userName, pass);
    }

    public void query(String sql) throws SQLException {
      try (Connection conn = getConnection(); Statement stmt = conn.createStatement(); ResultSet rs = stmt.executeQuery(sql)) {
        while (rs.next()) {
          int id = rs.getInt("req_seq_no");
          System.out.println(id);
        }
      }
    }

    /**
     * b.columnName = columnName; b.columnType = columnType; b.columnValue = columnValue;
     * 
     * 上記の値は使用可能です。
     */
    public List<B> queryAsObjects(String sql) {
      List<B> list = new ArrayList<B>();
      try (Connection conn = getConnection(); Statement stmt = conn.createStatement(); ResultSet rs = stmt.executeQuery(sql)) {
        ResultSetMetaData md = rs.getMetaData();
        int columns = md.getColumnCount();
        while (rs.next()) {
          for (int i = 1; i <= columns; i++) {
            B b = new B();
            String columnName = md.getColumnLabel(i);
            String columnType = md.getColumnTypeName(i);

            // 可在这里对日期、时间戳、BigDecimal 做定制格式化
            Object columnValue = rs.getObject(i);

            b.columnName = columnName;
            b.columnType = columnType;
            b.columnValue = columnValue;

            list.add(b);
          }

        }
      } catch (SQLException e) {
        e.printStackTrace();
      }
      return list;
    }

  }

  /**
   * Excel
   * 
   */
  class Excel {

    static B b;
    static Sub sub;
    static Application a;
    static ApplicationSub as;
    static Test my;

    static Util u = Sub.u;

    public void test() {}

  }

  /**
   * データBean
   */
  class B {

    static Excel e;
    static Oracle o;
    static Sub sub;

    public Map<String, List<String>> mapRows = new HashMap<String, List<String>>();
    // map

    public String key;
    public String value;

    // DB 操作

    // 物理名
    public String tableName;
    // 論理名
    public String tableComment;
    // 物理名
    public String columnName;
    // 論理名
    public String columnComment;
    // データ型
    public String columnType;
    // Not Null
    public String columnNotNull;
    // value
    public Object columnValue;

  }

  /**
   * Test
   */
  class Test {

    static Sub sub;
    static B b;

    static Oracle o = Sub.o;
    static Excel e = Sub.e;
    static Util u = Sub.u;
    static Application a = Sub.a;
    static ApplicationSub as = Sub.as;

    public void test() {

    }

  }

  /**
   * 固定のデータを入力する
   */
  public class Content {

    /**
     * 期待値の初期化
     */
    public String telegramExpectJson() {
      return """
          {
            "expectTestDataIDs": []
          }
             """;
    }


    /**
     * 期待値の初期化
     */
    public String telegramInputJson() {
      return """
          {
            "inputTestDataID": ""
          }
            """;
    }


    /**
     * データをクリアする
     */
    public String testClear() {

      String path = as.getTestDbTestCasePath();
      Map<File, String> map = u.findFileToMap(path);
      List<String> list = map.values().stream().toList();
      String r = JsonUtils.toJson(list).replace('[', '{').replace(']', '}');

      return String.format("""
          /**
           * データをクリアする
           */
          private void clear() {
            // 処理なし
            String[] sqls = %s;
            Statement stmt;
            try {
              stmt = DriverManager.getConnection(
                  env.getProperty("spring.datasource.url"),
                  env.getProperty("spring.datasource.username"),
                  env.getProperty("spring.datasource.password"))
                  .createStatement();
              for (String s : sqls) {
                stmt.addBatch(" delete from " + s);
              }
              stmt.executeBatch();
            } catch (SQLException e) {
              e.printStackTrace();
            }
          }
          """, r);

    }


  }

  /**
   * テンプレートクラス
   */
  class Template {
    /**
     * 
     */
    public String getSubTester() {
StringBuilder sb=new StringBuilder();
sb.append("////////////////////////////////////////////////////////////////////////////\n");
sb.append("//\n");
sb.append("// プログラム名：TOOL\n");
sb.append("// ファイル名：SubTester\n");
sb.append("// 作成者：徐楠\n");
sb.append("// 作成日：2025/12/16\n");
sb.append("// 備考：なし\n");
sb.append("// 変更履歴：なし\n");
sb.append("// 2025/12/16 徐楠 新規作成\n");
sb.append("//\n");
sb.append("////////////////////////////////////////////////////////////////////////////\n");
sb.append("\n");
sb.append("package " + Sub.class.getPackage().getName()+  ";\n");
sb.append("\n");
sb.append("import org.junit.jupiter.api.Tag;\n");
sb.append("import org.junit.jupiter.api.Test;\n");
sb.append("// import jp.or.jasdec.sbf.sb.ma.batch.biz.sub.SubSql;\n");
sb.append("\n");
sb.append("/**\n");
sb.append(" * 製造\n");
sb.append(" */\n");
sb.append("\n");
sb.append("@Tag(\"tool:製造\")\n");
sb.append("public class SubTester {\n");
sb.append("\n");
sb.append("  public static Sub.Util u = Sub.u;\n");
sb.append("  public static Sub.Application a = Sub.a;\n");
sb.append("  public static Sub.ApplicationSub as = Sub.as;\n");
sb.append("\n");
sb.append("  public static void main(String[] args) {\n");
sb.append("    // 从一条sql语句设置一个case\n");
sb.append("    // String caseNum = \"010005\";\n");
sb.append("    // Class<?> testClass = SBMAB002Tester.class;\n");
sb.append("    // as.setOneCaseFromSql(caseNum, testClass);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * SubIIn.txtからSubOut.txtに格納する\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void subTesterToTemplate() {\n");
sb.append("    as.subTesterToTemplate();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * Sub.javaのbak\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void subBak() {\n");
sb.append("    String[] fileNames = {\"JsonUtils\", \"KeyUI\", \"SubSql\", \"StartXUTester\", \"Index\", \"UTDbUtil\", \"ATest\", \"DbXlsmImportTool\"};\n");
sb.append("    as.subBak(fileNames);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 初始化 SubOut SubIn SubTester\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void testInit() {\n");
sb.append("    String[] fileNames = {\"ATest\", \"DbXlsmImportTool\"};\n");
sb.append("    as.testInit(fileNames);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 内容をSubIn.txtに入れる テスト\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void createKeyWord() {\n");
sb.append("    a.createKeyWord(\"at\");\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  // -------------------project\n");
sb.append("  /**\n");
sb.append("   * 创建mapper文件夹\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void createMapperDir() {\n");
sb.append("    as.createMapperDir();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 数据匹配\n");
sb.append("   * \n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void javaKeyPatterh() {\n");
sb.append("    String path = \"C:\\\\opt\\\\API\";\n");
sb.append("    String pattern = \"Empty(\";\n");
sb.append("    as.javaKeyPatterh(pattern, path);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  // -------------------project\n");
sb.append("\n");
sb.append("\n");
sb.append("  // ------------------------setOutputTestDbJsonValue------------------------\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 新しいテストを生成する\n");
sb.append("   * \n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void createNewTest() {\n");
sb.append("    // 【条件変更】\n");
sb.append("    String fromNum = \"010001\";\n");
sb.append("    String toNum = \"000015\";\n");
sb.append("    as.createNewTest(fromNum, toNum);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 获取表定义并打印到SubOut.json\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void getTableDefinePrint() {\n");
sb.append("    as.getTableDefinePrint(\"TB_FSBGWSECURITYNOTICELOG\");\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  // 1DB日志对齐使用 先生成DB模板json 通过xls\n");
sb.append("  // 2getJsonTemplateToAllDbTestCaseDir把模板json放到每一个case中\n");
sb.append("  // 3然后运行所有case得到log but was\n");
sb.append("  // 4setJsonDbExpectRealLength 把长度拉齐\n");
sb.append("  // 5然后运行所有case得到log row[i]\n");
sb.append("  // 6然后执行setJsonDbExpectRowI放入实际的值\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void getJsonTemplateToAllDbTestCaseDir() {\n");
sb.append("    // 选择一个模板放入所有的Db testCase目录中\n");
sb.append("    // as.getJsonTemplateToAllDbTestCaseDir(\"TB_FITTERMNOTICEINFO\", \"010\");\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void setJsonDbExpectRealLength() {\n");
sb.append("    // 更改数量\n");
sb.append("    as.setJsonDbExpectRealLength();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void setJsonDbExpectRowI() {\n");
sb.append("    // 修正数据\n");
sb.append("    as.setJsonDbExpectRowI();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void setJsonTelegramValueChange() {\n");
sb.append("    // 電文データ修正\n");
sb.append("    as.setJsonTelegramValueChange();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void setTestMessage() {\n");
sb.append("    // エラーメッセージを設定する\n");
sb.append("    as.setTestMessage();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  // ------------------------------telegram\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト Cannot invoke\n");
sb.append("   * \"jp.co.skywave.ut.telegram.bean.InterfaceExpectTestDataSet.getExpectTestDataIDs()\" because\n");
sb.append("   * \"expectSet\" is null\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void createInitTelegramExpectInputSet() {\n");
sb.append("    // 初始化创建电文期望值文件\n");
sb.append("    String caseNum = \"111115\";\n");
sb.append("    as.createInitTelegramExpectInputSet(caseNum);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void findJsonTelegramOK() {\n");
sb.append("    // 找出匹配电文\n");
sb.append("    as.findJsonTelegramOK();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void createJsonTelegramOK() {\n");
sb.append("    // 生成匹配电文数据\n");
sb.append("    String caseNum = \"111115\";\n");
sb.append("    as.createJsonTelegramOK(caseNum);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  // ------------------------------telegram\n");
sb.append("\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void setJsonDbInputValueChange() {\n");
sb.append("    // 自分でinputのjsonデータを設定する\n");
sb.append("    String testNum = \"010105\";\n");
sb.append("    // String param = \"TB_FITISSUERNXT \\\"CO_CD\\\": NULL,\";\n");
sb.append("    String param = \"TB_FSBACCOUNT \\\"CO_CD\\\": \\\"1\\\",\";\n");
sb.append("    // as.setJsonDbInputValueChange(testNum, param);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void sqlAsAndBeanFieldCompare() {\n");
sb.append("    // sql别名和bean字段一致性检测 → SubOut.log\n");
sb.append("    as.sqlAsAndBeanFieldCompare();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * → SubOut.txt\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void testClearContent() {\n");
sb.append("    as.testClearContent();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  // ------------------------setOutputTestDbJsonValue------------------------\n");
sb.append("\n");
sb.append("  // ------------------------iput test Data------------------------\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト 1\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void jsonDataCheck() {\n");
sb.append("    // 放入json中合适数据使得sql查询成立 检测数据是否完成在SubOut.json,SubOut.log中查看\n");
sb.append("\n");
sb.append("    // 当前测试号\n");
sb.append("    String testNum = \"000015\";\n");
sb.append("    // SubSql.jsonDataCheck(testNum, true);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト 2\n");
sb.append("   * \n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void jsonDataInit() throws Exception {\n");
sb.append("    // SubOut.log 文件会递增 手动先删除内容\n");
sb.append("    // SubOUt.json 解析出的sql结构\n");
sb.append("    // 初始化DB数据 来源 C:\\zop\\testJsonCaseData\n");
sb.append("\n");
sb.append("    // 当前测试号\n");
sb.append("    String testNum = \"000015\";\n");
sb.append("    // SubSql.jsonDataInit(testNum);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テスト 3\n");
sb.append("   */\n");
sb.append("  @SuppressWarnings(\"static-access\")\n");
sb.append("  @Test\n");
sb.append("  public void jsonDataSet() {\n");
sb.append("    // 放入json中合适数据使得sql查询成立 数据在SubOut.json,SubOut.log中查看\n");
sb.append("\n");
sb.append("    // 当前测试号\n");
sb.append("    String testNum = \"000015\";\n");
sb.append("    // SubSql.jsonDataSet(testNum, false);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  // ------------------------iput test Data------------------------\n");
sb.append("\n");
sb.append("}\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
return sb.toString();
    }
    /**
     * 
     */
    public String getJsonUtils() {
StringBuilder sb=new StringBuilder();
sb.append("////////////////////////////////////////////////////////////////////////////\n");
sb.append("//\n");
sb.append("// プログラム名：JsonUtils\n");
sb.append("// ファイル名：JsonUtils\n");
sb.append("// 作成者：automatic generator\n");
sb.append("// 作成日：2026/01/14\n");
sb.append("// 備考：\n");
sb.append("//\n");
sb.append("////////////////////////////////////////////////////////////////////////////\n");
sb.append("\n");
sb.append("package " + Sub.class.getPackage().getName()+  ".sub;\n");
sb.append("\n");
sb.append("import java.io.IOException;\n");
sb.append("import java.text.DateFormat;\n");
sb.append("import java.text.SimpleDateFormat;\n");
sb.append("\n");
sb.append("import com.fasterxml.jackson.annotation.JsonInclude;\n");
sb.append("import com.fasterxml.jackson.core.type.TypeReference;\n");
sb.append("import com.fasterxml.jackson.databind.ObjectMapper;\n");
sb.append("import org.slf4j.Logger;\n");
sb.append("import org.slf4j.LoggerFactory;\n");
sb.append("\n");
sb.append("\n");
sb.append("/**\n");
sb.append(" * JSON形式のパース\n");
sb.append(" */\n");
sb.append("public class JsonUtils {\n");
sb.append("\n");
sb.append("  private static final Logger logger = LoggerFactory.getLogger(JsonUtils.class);\n");
sb.append("\n");
sb.append("  // からのフィールドをむしする\n");
sb.append("  private static ObjectMapper OBJECT_MAPPER = new ObjectMapper();\n");
sb.append("\n");
sb.append("  private static ObjectMapper OBJECT_MAPPER_EXT = new ObjectMapper();\n");
sb.append("\n");
sb.append("  static {\n");
sb.append("\n");
sb.append("    DateFormat df = new SimpleDateFormat(\"yyyy/MM/dd HH:mm:ss\");\n");
sb.append("    // からちとからのコレクションをむしす\n");
sb.append("    // OBJECT_MAPPER.setSerializationInclusion(JsonInclude.Include.NON_EMPTY);\n");
sb.append("    OBJECT_MAPPER.setDateFormat(df);\n");
sb.append("\n");
sb.append("    // 値が NULL の場合、属性値は返却されません。\n");
sb.append("    OBJECT_MAPPER_EXT.setSerializationInclusion(JsonInclude.Include.NON_NULL);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("  * \n");
sb.append("  */\n");
sb.append("  public String toJson(Object object) {\n");
sb.append("    try {\n");
sb.append("      String json = OBJECT_MAPPER.writeValueAsString(object);\n");
sb.append("      return json;\n");
sb.append("    } catch (Exception e) {\n");
sb.append("      throw new RuntimeException(\"Json.format error:\" + object, e);\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 対象が NULL の場合は表示されません。コレクションが空の場合は「[]」を返します。\n");
sb.append("   */\n");
sb.append("  public String toJsonExt(Object object) {\n");
sb.append("    try {\n");
sb.append("      return OBJECT_MAPPER_EXT.writeValueAsString(object);\n");
sb.append("    } catch (IOException e) {\n");
sb.append("      logger.warn(\"write to json string error:\" + object, e);\n");
sb.append("    }\n");
sb.append("    return null;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("  * \n");
sb.append("  */\n");
sb.append("  public <T> T fromJson(String jsonString, Class<T> clazz) {\n");
sb.append("    if (jsonString == null) {\n");
sb.append("      return null;\n");
sb.append("    }\n");
sb.append("    try {\n");
sb.append("      return OBJECT_MAPPER.readValue(jsonString, clazz);\n");
sb.append("    } catch (IOException e) {\n");
sb.append("      throw new RuntimeException(\"parse json string error:\" + jsonString, e);\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * \n");
sb.append("   */\n");
sb.append("  public <T> T fromJson(String jsonString, TypeReference<T> typeReference) {\n");
sb.append("    if (jsonString == null) {\n");
sb.append("      return null;\n");
sb.append("    }\n");
sb.append("    try {\n");
sb.append("      return OBJECT_MAPPER.readValue(jsonString, typeReference);\n");
sb.append("    } catch (IOException e) {\n");
sb.append("      throw new RuntimeException(\"parse json string error:\" + jsonString, e);\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("}\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
return sb.toString();
    }
    /**
     * 
     */
    public String getKeyUI() {
StringBuilder sb=new StringBuilder();
sb.append("////////////////////////////////////////////////////////////////////////////\n");
sb.append("//\n");
sb.append("// プログラム名：Sub\n");
sb.append("// ファイル名：Sub\n");
sb.append("// 作成者：automatic generator\n");
sb.append("// 作成日：2026/01/14\n");
sb.append("// 備考：\n");
sb.append("//\n");
sb.append("////////////////////////////////////////////////////////////////////////////\n");
sb.append("\n");
sb.append("package " + Sub.class.getPackage().getName()+  ".sub;\n");
sb.append("\n");
sb.append("import org.eclipse.swt.SWT;\n");
sb.append("import org.eclipse.swt.custom.StyledText;\n");
sb.append("import org.eclipse.swt.events.SelectionAdapter;\n");
sb.append("import org.eclipse.swt.events.SelectionEvent;\n");
sb.append("import org.eclipse.swt.graphics.Point;\n");
sb.append("import org.eclipse.swt.graphics.Rectangle;\n");
sb.append("import org.eclipse.swt.layout.FillLayout;\n");
sb.append("import org.eclipse.swt.layout.GridData;\n");
sb.append("import org.eclipse.swt.layout.GridLayout;\n");
sb.append("import org.eclipse.swt.widgets.Button;\n");
sb.append("import org.eclipse.swt.widgets.Composite;\n");
sb.append("import org.eclipse.swt.widgets.Control;\n");
sb.append("import org.eclipse.swt.widgets.Display;\n");
sb.append("import org.eclipse.swt.widgets.Label;\n");
sb.append("import org.eclipse.swt.widgets.MessageBox;\n");
sb.append("import org.eclipse.swt.widgets.Shell;\n");
sb.append("\n");
sb.append("/**\n");
sb.append(" * \n");
sb.append(" */\n");
sb.append("public class KeyUI {\n");
sb.append("\n");
sb.append("  protected Shell shell;\n");
sb.append("  protected Display display;\n");
sb.append("  private Composite composite;\n");
sb.append("\n");
sb.append("  // 文本显示面板\n");
sb.append("  private StyledText styledText;\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * \n");
sb.append("   */\n");
sb.append("  public static void main(String[] args) {\n");
sb.append("    try {\n");
sb.append("      KeyUI window = new KeyUI();\n");
sb.append("      window.open();\n");
sb.append("    } catch (Exception e) {\n");
sb.append("      e.printStackTrace();\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * Open the window.\n");
sb.append("   */\n");
sb.append("  public void open() {\n");
sb.append("\n");
sb.append("    display = Display.getDefault();\n");
sb.append("    createContents();\n");
sb.append("\n");
sb.append("    // 放在右下角\n");
sb.append("    Rectangle screen = display.getPrimaryMonitor().getBounds();\n");
sb.append("    Rectangle rect = shell.getBounds();\n");
sb.append("    int x = screen.x + screen.width - rect.width;\n");
sb.append("    int y = screen.y + screen.height - rect.height;\n");
sb.append("    shell.setLocation(x, y);\n");
sb.append("\n");
sb.append("    new Label(composite, SWT.NONE);\n");
sb.append("    new Label(composite, SWT.NONE);\n");
sb.append("    new Label(composite, SWT.NONE);\n");
sb.append("    new Label(composite, SWT.NONE);\n");
sb.append("    new Label(composite, SWT.NONE);\n");
sb.append("    new Label(composite, SWT.NONE);\n");
sb.append("    new Label(composite, SWT.NONE);\n");
sb.append("    new Label(composite, SWT.NONE);\n");
sb.append("\n");
sb.append("    // 设置窗体透明\n");
sb.append("    shell.setAlpha(180); // 0~255\n");
sb.append("\n");
sb.append("    shell.open();\n");
sb.append("\n");
sb.append("    shell.layout();\n");
sb.append("\n");
sb.append("    while (!shell.isDisposed()) {\n");
sb.append("      if (!display.readAndDispatch()) {\n");
sb.append("        display.sleep();\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * Create contents of the window.\n");
sb.append("   */\n");
sb.append("  protected void createContents() {\n");
sb.append("\n");
sb.append("    shell = new Shell(display, SWT.RESIZE | SWT.ON_TOP);\n");
sb.append("\n");
sb.append("    shell.setSize(479, 300);\n");
sb.append("    shell.setLayout(new FillLayout(SWT.HORIZONTAL));\n");
sb.append("\n");
sb.append("    // 总的三行布局\n");
sb.append("    composite = new Composite(shell, SWT.NONE);\n");
sb.append("    composite.setBackground(shell.getDisplay().getSystemColor(SWT.COLOR_WHITE));\n");
sb.append("    composite.setLayout(new GridLayout(11, false));\n");
sb.append("\n");
sb.append("    addDragSupport(shell);\n");
sb.append("    addDragSupport(composite);\n");
sb.append("\n");
sb.append("    // 面板布局\n");
sb.append("    styledText = new StyledText(composite, SWT.FULL_SELECTION | SWT.WRAP | SWT.H_SCROLL | SWT.V_SCROLL);\n");
sb.append("    GridData gridData = new GridData(SWT.FILL, SWT.FILL, true, true, 11, 1);\n");
sb.append("    gridData.widthHint = 669;\n");
sb.append("    styledText.setLayoutData(gridData);\n");
sb.append("\n");
sb.append("    createButton(\"啊\");\n");
sb.append("\n");
sb.append("    createButton(\"吧\");\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   */\n");
sb.append("  protected Button createButton(String text) {\n");
sb.append("    Button button = new Button(composite, SWT.NONE);\n");
sb.append("    button.setText(text);\n");
sb.append("    button.addSelectionListener(new SelectionAdapter() {\n");
sb.append("      @Override\n");
sb.append("      public void widgetSelected(SelectionEvent arg0) {\n");
sb.append("        MessageBox mb = new MessageBox(shell, SWT.OK | SWT.ICON_INFORMATION);\n");
sb.append("        mb.setText(\"提示\");\n");
sb.append("        mb.setMessage(\"操作成功！\");\n");
sb.append("        mb.open();\n");
sb.append("      }\n");
sb.append("    });\n");
sb.append("    return button;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   */\n");
sb.append("  private void addDragSupport(Control ctrl) {\n");
sb.append("    ctrl.addListener(SWT.MouseDown, e -> ctrl.setData(\"drag\", new Point(e.x, e.y)));\n");
sb.append("\n");
sb.append("    ctrl.addListener(SWT.MouseMove, e -> {\n");
sb.append("      Point p = (Point) ctrl.getData(\"drag\");\n");
sb.append("      if (p != null && (e.stateMask & SWT.BUTTON1) != 0) {\n");
sb.append("        Point loc = shell.getLocation();\n");
sb.append("        shell.setLocation(loc.x + e.x - p.x, loc.y + e.y - p.y);\n");
sb.append("      }\n");
sb.append("    });\n");
sb.append("\n");
sb.append("    ctrl.addListener(SWT.MouseUp, e -> ctrl.setData(\"drag\", null));\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * \n");
sb.append("   */\n");
sb.append("  class KeyUISub {\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("}\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
return sb.toString();
    }
    /**
     * 
     */
    public String getSubSql() {
StringBuilder sb=new StringBuilder();
sb.append("////////////////////////////////////////////////////////////////////////////\n");
sb.append("//\n");
sb.append("// プログラム名：Sub\n");
sb.append("// ファイル名：Sub\n");
sb.append("// 作成者：automatic generator\n");
sb.append("// 作成日：2026/01/14\n");
sb.append("// 備考：implementation 'com.github.jsqlparser:jsqlparser:4.6'\n");
sb.append("//\n");
sb.append("////////////////////////////////////////////////////////////////////////////\n");
sb.append("\n");
sb.append("package " + Sub.class.getPackage().getName()+  ".sub;\n");
sb.append("\n");
sb.append("import java.io.File;\n");
sb.append("import java.util.ArrayList;\n");
sb.append("import java.util.HashMap;\n");
sb.append("import java.util.LinkedHashMap;\n");
sb.append("import java.util.List;\n");
sb.append("import java.util.Map;\n");
sb.append("import java.util.regex.Matcher;\n");
sb.append("import java.util.regex.Pattern;\n");
sb.append("import org.junit.platform.engine.discovery.DiscoverySelectors;\n");
sb.append("import org.junit.platform.launcher.Launcher;\n");
sb.append("import org.junit.platform.launcher.LauncherDiscoveryRequest;\n");
sb.append("import org.junit.platform.launcher.core.LauncherDiscoveryRequestBuilder;\n");
sb.append("import org.junit.platform.launcher.core.LauncherFactory;\n");
sb.append("import org.junit.platform.launcher.listeners.SummaryGeneratingListener;\n");
sb.append("import org.junit.platform.launcher.listeners.TestExecutionSummary;\n");
sb.append("import org.slf4j.LoggerFactory;\n");
sb.append("import ch.qos.logback.classic.LoggerContext;\n");
sb.append("import ch.qos.logback.classic.joran.JoranConfigurator;\n");
sb.append("import jp.or.jasdec.sbf.sb.ma.batch.biz.Sub;\n");
sb.append("import net.sf.jsqlparser.JSQLParserException;\n");
sb.append("import net.sf.jsqlparser.expression.Expression;\n");
sb.append("import net.sf.jsqlparser.expression.Function;\n");
sb.append("import net.sf.jsqlparser.expression.operators.conditional.AndExpression;\n");
sb.append("import net.sf.jsqlparser.expression.operators.conditional.OrExpression;\n");
sb.append("import net.sf.jsqlparser.expression.operators.relational.EqualsTo;\n");
sb.append("import net.sf.jsqlparser.expression.operators.relational.ExpressionList;\n");
sb.append("import net.sf.jsqlparser.expression.operators.relational.InExpression;\n");
sb.append("import net.sf.jsqlparser.parser.CCJSqlParserUtil;\n");
sb.append("import net.sf.jsqlparser.schema.Table;\n");
sb.append("import net.sf.jsqlparser.statement.Statement;\n");
sb.append("import net.sf.jsqlparser.statement.select.Join;\n");
sb.append("import net.sf.jsqlparser.statement.select.PlainSelect;\n");
sb.append("import net.sf.jsqlparser.statement.select.Select;\n");
sb.append("import net.sf.jsqlparser.statement.select.SelectVisitorAdapter;\n");
sb.append("\n");
sb.append("/**\n");
sb.append(" * \n");
sb.append(" */\n");
sb.append("public class SubSql {\n");
sb.append("  // implementation 'com.github.jsqlparser:jsqlparser:4.6'\n");
sb.append("  // implementation 'org.junit.platform:junit-platform-launcher:1.9.1'\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 主函数（测试）\n");
sb.append("   */\n");
sb.append("  public static void main(String[] args) {\n");
sb.append("\n");
sb.append("    // 初始化 input json 数据存储在 C:\\zop\\testJsonCaseData\n");
sb.append("    // jsonDataInit(\"111111\");\n");
sb.append("\n");
sb.append("    // 检测数据开关打开 true 打印数据可查看是否存在没有完成解析的数据存在\n");
sb.append("    // jsonDataCheck(\"111111\", true);\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  // -------------------------------------------------测试一个方法\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 从一条sql语句设置一个case\n");
sb.append("   */\n");
sb.append("  public static void setOneCaseFromSql(String caseNum, Class<?> testClass) {\n");
sb.append("    // SubIn.txt放入sql语句 此语句经过了 jsonDataCheck() 保证语句正确后处理\n");
sb.append("    try {\n");
sb.append("\n");
sb.append("      // // 初始化文件到 input output\n");
sb.append("      // SubSql.jsonDataInit(caseNum);\n");
sb.append("      // // 按照sql的关联 保证数据整合性,使得数据可以查询出来\n");
sb.append("      // SubSql.jsonDataSet(caseNum, false);\n");
sb.append("\n");
sb.append("      String methodName = \"test\" + caseNum;\n");
sb.append("      // 执行一次放入的语句 查看log日志中是否存在问题\n");
sb.append("      testOneMethod(caseNum, methodName, testClass, 0);\n");
sb.append("    } catch (Exception e) {\n");
sb.append("      e.printStackTrace();\n");
sb.append("    }\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 对一个测试方法进行测试\n");
sb.append("   */\n");
sb.append("  public static void testOneMethod(String caseNum, String methodName, Class<?> testClass, Integer flag) {\n");
sb.append("\n");
sb.append("    SummaryGeneratingListener summaryListener = new SummaryGeneratingListener();\n");
sb.append("\n");
sb.append("    LauncherDiscoveryRequest request =\n");
sb.append("        LauncherDiscoveryRequestBuilder.request().selectors(DiscoverySelectors.selectMethod(testClass, methodName)).build();\n");
sb.append("\n");
sb.append("    Launcher launcher = LauncherFactory.create();\n");
sb.append("    launcher.registerTestExecutionListeners(summaryListener);\n");
sb.append("    launcher.execute(request);\n");
sb.append("\n");
sb.append("    TestExecutionSummary summary = summaryListener.getSummary();\n");
sb.append("    long failures = summary.getFailures().size();\n");
sb.append("\n");
sb.append("    if (failures == 0) {\n");
sb.append("      System.out.println(\"=== \" + methodName + \"测试成功 ===\");\n");
sb.append("    } else {\n");
sb.append("      System.out.println(\"=== 测试失败 \" + failures + \" 个 ===\");\n");
sb.append("      String path = Sub.u.getProjectDir(\"log\") + \"app.log\";\n");
sb.append("      String content = Sub.u.readFile(path);\n");
sb.append("\n");
sb.append("      if (Sub.u.isContentMatch(content, \"because \\\"expectSet\\\" is null\")) {\n");
sb.append("        // 说明电文没有文件需要创建新的文件\n");
sb.append("        Sub.as.createInitTelegramExpectInputSet(caseNum);\n");
sb.append("        Sub.as.clearLog();\n");
sb.append("        forceResetLogback();\n");
sb.append("        // 重新执行一次\n");
sb.append("        testOneMethod(caseNum, methodName, testClass, 0);\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // 找出匹配电文\n");
sb.append("      if (Sub.u.isContentMatch(content, \"【差異】NotDefine|【差異】ChannelDiff|NG:Length\")) {\n");
sb.append("        // 说明电文没有文件需要创建新的文件\n");
sb.append("        Sub.as.findJsonTelegramOK();\n");
sb.append("        Sub.as.clearLog();\n");
sb.append("        forceResetLogback();\n");
sb.append("        // 重新执行一次\n");
sb.append("        testOneMethod(caseNum, methodName, testClass, 0);\n");
sb.append("      }\n");
sb.append("      // ------上面完成文件的初始化\n");
sb.append("\n");
sb.append("      // 下面开始放入数据\n");
sb.append("      // 查看电文是否有差分\n");
sb.append("      if (Sub.as.isTelegramValueChange()) {\n");
sb.append("        if (flag == 0) {\n");
sb.append("          // 初始化自己的电文\n");
sb.append("          Sub.as.createJsonTelegramOK(caseNum);\n");
sb.append("          Sub.as.clearLog();\n");
sb.append("          forceResetLogback();\n");
sb.append("          testOneMethod(caseNum, methodName, testClass, 1);\n");
sb.append("        } else {\n");
sb.append("          // flag为1说明初始化完成需要放入完整数据了\n");
sb.append("          Sub.as.setJsonTelegramValueChange();\n");
sb.append("          Sub.as.clearLog();\n");
sb.append("          forceResetLogback();\n");
sb.append("          // 重新执行一次\n");
sb.append("          testOneMethod(caseNum, methodName, testClass, 1);\n");
sb.append("        }\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      flag = 1;\n");
sb.append("\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 重新配置log防止写文件时总是原来的数据\n");
sb.append("   */\n");
sb.append("  public static void forceResetLogback() {\n");
sb.append("\n");
sb.append("    LoggerContext context = (LoggerContext) LoggerFactory.getILoggerFactory();\n");
sb.append("    try {\n");
sb.append("      // 1. 停止 logback\n");
sb.append("      context.stop();\n");
sb.append("\n");
sb.append("      // 2. 读取 logback.xml 配置重新加载\n");
sb.append("      JoranConfigurator configurator = new JoranConfigurator();\n");
sb.append("      configurator.setContext(context);\n");
sb.append("\n");
sb.append("      // 清理旧配置\n");
sb.append("      context.reset();\n");
sb.append("\n");
sb.append("      // ⭐⭐ 使用你的 logback.xml 路径\n");
sb.append("      configurator.doConfigure(Sub.as.getSrcTestResourcesPath() + \"logback.xml\");\n");
sb.append("\n");
sb.append("      System.out.println(\"Logback 已重新加载配置\");\n");
sb.append("    } catch (Exception e) {\n");
sb.append("      e.printStackTrace();\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("  // -------------------------------------------------测试一个方法\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 放入json中合适数据使得sql查询成立\n");
sb.append("   */\n");
sb.append("  public static void jsonDataSet(String toTestNum, Boolean isTestOpen) {\n");
sb.append("    String json = getJson();\n");
sb.append("    List<String> changeList = getWhereOrOn(json, isTestOpen);\n");
sb.append("\n");
sb.append("    for (String content : changeList) {\n");
sb.append("      // String param = \"TB_FITISSUERNXT \\\"CO_CD\\\": NULL,\";\n");
sb.append("      // String param = \"TB_FITISSUERNXT \\\"CO_CD\\\": \\\"1\\\",\";\n");
sb.append("      String param = content;\n");
sb.append("      Sub.as.setJsonDbInputValueChange(toTestNum, param);\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 放入json中合适数据使得sql查询成立 检测数据是否完成在SubOut.json中查看\n");
sb.append("   * \n");
sb.append("   * <pre>\n");
sb.append("   * 检测\n");
sb.append("   * </pre>\n");
sb.append("   */\n");
sb.append("  public static void jsonDataCheck(String toTestNum, Boolean isTestOpen) {\n");
sb.append("    String json = getJson();\n");
sb.append("    getWhereOrOn(json, isTestOpen);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 初始化测试数据 sql来自SubIn.txt经过转换中间放入SubOut.Json\n");
sb.append("   */\n");
sb.append("  public static void jsonDataInit(String toTestNum) throws Exception {\n");
sb.append("    String json = getJson();\n");
sb.append("    jsonTableToTestInputDir(json, toTestNum);\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * \"TB_FSBREQUEST.RT_STAT_KBN = '1'\",\n");
sb.append("   * \n");
sb.append("   * \"TB_FSBREQUEST.MSG_1_CHR#7,8# =TB_FSBGWSECURITYNOTICELOG.SEC_ACPT_NO\n");
sb.append("   * \n");
sb.append("   * 比较差异 日志输出到SubOut.log\n");
sb.append("   */\n");
sb.append("  public static List<String> getWhereOrOn(String json, Boolean isTestOpen) {\n");
sb.append("\n");
sb.append("    StringBuilder logSb = new StringBuilder();\n");
sb.append("\n");
sb.append("    // json 中需要改变值的数据\n");
sb.append("    List<String> changeList = new ArrayList<String>();\n");
sb.append("\n");
sb.append("    // 1. 解析 JSON 为 Map\n");
sb.append("    Map<String, Object> root = Sub.JsonUtils.fromJson(json, new com.fasterxml.jackson.core.type.TypeReference<Map<String, Object>>() {});\n");
sb.append("\n");
sb.append("    // 3. 取 whereOrOn\n");
sb.append("    @SuppressWarnings(\"unchecked\")\n");
sb.append("    List<String> whereOrOn = (List<String>) root.get(\"whereOrOn\");\n");
sb.append("\n");
sb.append("    // 存储已经知道的数据,后面如果用到了这个字段就先获取这里面的数据 否则会出现数据不一致的情况\n");
sb.append("    // 如前面 kbn=65 后面的数据也要用这个值\n");
sb.append("    Map<String, String> mapValue = new HashMap<String, String>();\n");
sb.append("\n");
sb.append("    // TB_FITISSUERNXT \"CO_CD\": NULL,\n");
sb.append("    // TB_FITISSUERNXT \"CO_CD\": \"1\",\n");
sb.append("    for (String condition : whereOrOn) {\n");
sb.append("\n");
sb.append("      if (isTestOpen) {\n");
sb.append("        String message = \"================\" + condition;\n");
sb.append("        System.out.println(message);\n");
sb.append("        logSb.append(message + \"\\n\");\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // 情况1 TB_FSBREQUEST.SELLER_USR_CD#1,5# = TB_FSBCORP.ACCMG_CD\n");
sb.append("      Pattern p = Pattern.compile(\"([A-Z0-9_]+)\\\\.([A-Z0-9_]+)#(\\\\d+),(\\\\d+)#\\\\s*=\\\\s*([A-Z0-9_]+)\\\\.([A-Z0-9_]+)\");\n");
sb.append("      Matcher m = p.matcher(condition);\n");
sb.append("\n");
sb.append("      if (m.find()) {\n");
sb.append("\n");
sb.append("        String leftTable = m.group(1); // TB_FSBREQUEST\n");
sb.append("        String leftField = m.group(2); // SELLER_USR_CD\n");
sb.append("        int start = Integer.parseInt(m.group(3)); // 6\n");
sb.append("        int len = Integer.parseInt(m.group(4)); // 2 从第六个开始取两位\n");
sb.append("\n");
sb.append("        String rightTable = m.group(5); // TB_FSBOPENACC\n");
sb.append("        String rightField = m.group(6); // KBN_ACC_CD\n");
sb.append("\n");
sb.append("        // 根据你的规则：原始长度 = start + len\n");
sb.append("        int originalLength = start + len;\n");
sb.append("\n");
sb.append("        // 左侧的完整值用 ? 占位\n");
sb.append("        String leftValue = getTestValue(originalLength, null, null);\n");
sb.append("        leftValue = getStoreValue(mapValue, leftTable, leftField, leftValue);\n");
sb.append("\n");
sb.append("        // 右侧值 = 截取部分（用 * 占位）\n");
sb.append("        String rightValue = getTestValue(originalLength, start, len);\n");
sb.append("        rightValue = getStoreValue(mapValue, rightTable, rightField, rightValue);\n");
sb.append("\n");
sb.append("        String leftString = leftTable + \" \\\"\" + leftField + \"\\\": \\\"\" + leftValue + \"\\\",\";\n");
sb.append("        String rightString = rightTable + \" \\\"\" + rightField + \"\\\": \\\"\" + rightValue + \"\\\",\";\n");
sb.append("        changeList.add(leftString);\n");
sb.append("        changeList.add(rightString);\n");
sb.append("        if (isTestOpen) {\n");
sb.append("          // 输出\n");
sb.append("          System.out.println(leftString);\n");
sb.append("          logSb.append(leftString + \"\\n\");\n");
sb.append("          System.out.println(rightString);\n");
sb.append("          logSb.append(rightString + \"\\n\");\n");
sb.append("        }\n");
sb.append("        continue;\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // 情况2 TB_FSBREQUEST.RT_STAT_KBN = '1'\n");
sb.append("      p = Pattern.compile(\"([A-Z0-9_]+)\\\\.([A-Z0-9_#]+)\\\\s*=\\\\s*'([^']*)'\");\n");
sb.append("      m = p.matcher(condition);\n");
sb.append("      if (m.find()) {\n");
sb.append("        String table = m.group(1); // 表字段\n");
sb.append("        String field = m.group(2); // 字段\n");
sb.append("        String value = m.group(3); // 值\n");
sb.append("\n");
sb.append("        value = getStoreValue(mapValue, table, field, value);\n");
sb.append("\n");
sb.append("        String changeString = \"\" + table + \" \\\"\" + field + \"\\\": \\\"\" + value + \"\\\",\";\n");
sb.append("        changeList.add(changeString);\n");
sb.append("        if (isTestOpen) {\n");
sb.append("          System.out.println(changeString);\n");
sb.append("          logSb.append(changeString + \"\\n\");\n");
sb.append("        }\n");
sb.append("        continue;\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // 情况3 TB_FSBREQUEST.PRG_STAT_KBN IN (65,6D)\n");
sb.append("      p = Pattern.compile(\"([A-Z0-9_]+)\\\\.([A-Z0-9_]+)\\\\s+IN\\\\s*\\\\(([^)]+)\\\\)\");\n");
sb.append("      m = p.matcher(condition);\n");
sb.append("\n");
sb.append("      if (m.find()) {\n");
sb.append("        String table = m.group(1); // TB_FSBREQUEST\n");
sb.append("        String field = m.group(2); // PRG_STAT_KBN\n");
sb.append("        String list = m.group(3); // 65,6D\n");
sb.append("\n");
sb.append("        // 拿到逗号前面的第一项\n");
sb.append("        String firstValue = list.split(\",\")[0].trim();\n");
sb.append("        firstValue = getStoreValue(mapValue, table, field, firstValue);\n");
sb.append("\n");
sb.append("        String changeString = table + \" \\\"\" + field + \"\\\": \\\"\" + firstValue + \"\\\",\";\n");
sb.append("        changeList.add(changeString);\n");
sb.append("        if (isTestOpen) {\n");
sb.append("          System.out.println(changeString);\n");
sb.append("          logSb.append(changeString + \"\\n\");\n");
sb.append("        }\n");
sb.append("        continue;\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // 情况4 TB_FSBSUBORGNI.CO_CD = TB_FSBBANK.CO_CD\n");
sb.append("      p = Pattern.compile(\"([A-Z0-9_]+)\\\\.([A-Z0-9_]+)\\\\s*=\\\\s*([A-Z0-9_]+)\\\\.([A-Z0-9_]+)\");\n");
sb.append("\n");
sb.append("      m = p.matcher(condition);\n");
sb.append("\n");
sb.append("      if (m.find()) {\n");
sb.append("\n");
sb.append("        String leftTable = m.group(1); // TB_FSBREQUEST\n");
sb.append("        String leftField = m.group(2); // ISIN_CD\n");
sb.append("        String rightTable = m.group(3); // TB_FSBREQREC\n");
sb.append("        String rightField = m.group(4); // ISIN_CD\n");
sb.append("        String leftvalue = \"1\";\n");
sb.append("        String rightvalue = \"1\";\n");
sb.append("\n");
sb.append("        String leftkey = leftTable + \"_\" + leftField;\n");
sb.append("        String rightkey = rightTable + \"_\" + rightField;\n");
sb.append("\n");
sb.append("        if (mapValue.get(leftkey) != null && mapValue.get(rightkey) == null) {\n");
sb.append("          leftvalue = mapValue.get(leftkey);\n");
sb.append("          rightvalue = leftvalue;\n");
sb.append("        }\n");
sb.append("\n");
sb.append("        if (mapValue.get(leftkey) == null && mapValue.get(rightkey) != null) {\n");
sb.append("          leftvalue = mapValue.get(rightkey);\n");
sb.append("          rightvalue = leftvalue;\n");
sb.append("        }\n");
sb.append("\n");
sb.append("        String leftString = leftTable + \" \\\"\" + leftField + \"\\\": \\\"\" + leftvalue + \"\\\",\";\n");
sb.append("        String rightString = rightTable + \" \\\"\" + rightField + \"\\\": \\\"\" + rightvalue + \"\\\",\";\n");
sb.append("\n");
sb.append("        changeList.add(leftString);\n");
sb.append("        changeList.add(rightString);\n");
sb.append("\n");
sb.append("        if (isTestOpen) {\n");
sb.append("          System.out.println(leftString);\n");
sb.append("          logSb.append(leftString + \"\\n\");\n");
sb.append("          System.out.println(rightString);\n");
sb.append("          logSb.append(rightString + \"\\n\");\n");
sb.append("        }\n");
sb.append("        continue;\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // 这种情况是单表查询\n");
sb.append("      // 情况5: 字段 = '常量' 例如: REQ_EXIST_FLG = '0'\n");
sb.append("      p = Pattern.compile(\"([A-Z0-9_#]+)\\\\s*=\\\\s*'([^']*)'\");\n");
sb.append("\n");
sb.append("      m = p.matcher(condition);\n");
sb.append("\n");
sb.append("      if (m.find()) {\n");
sb.append("\n");
sb.append("        String table = m.group(1); // 可能是表名，也可能是字段名（无表）\n");
sb.append("\n");
sb.append("        @SuppressWarnings(\"unchecked\")\n");
sb.append("        Map<String, String> aliasMapping = (Map<String, String>) root.get(\"aliasMapping\");\n");
sb.append("        for (Map.Entry<String, String> entry : aliasMapping.entrySet()) {\n");
sb.append("          // String alias = entry.getKey();\n");
sb.append("          table = entry.getValue();\n");
sb.append("        }\n");
sb.append("\n");
sb.append("        String field = m.group(1); // 字段名（如果有表名）\n");
sb.append("        String value = m.group(2); // '0' 中的 0\n");
sb.append("\n");
sb.append("        String jsonString = table + \" \\\"\" + field + \"\\\": \\\"\" + value + \"\\\",\";\n");
sb.append("\n");
sb.append("        changeList.add(jsonString);\n");
sb.append("\n");
sb.append("        if (isTestOpen) {\n");
sb.append("          System.out.println(jsonString);\n");
sb.append("          logSb.append(jsonString + \"\\n\");\n");
sb.append("        }\n");
sb.append("        continue;\n");
sb.append("      }\n");
sb.append("\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    Sub.u.outFile(logSb.toString(), Sub.as.getSubOutLogPath());\n");
sb.append("    return changeList;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 存储值 如果存在则下面的数据使用存储值\n");
sb.append("   */\n");
sb.append("  private static String getStoreValue(Map<String, String> mapValue, String table, String field, String value) {\n");
sb.append("    String key = table + \"_\" + field;\n");
sb.append("    if (mapValue.get(key) == null) {\n");
sb.append("      mapValue.put(key, value);\n");
sb.append("    } else {\n");
sb.append("      value = mapValue.get(key);\n");
sb.append("    }\n");
sb.append("    return value;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 测试值\n");
sb.append("   */\n");
sb.append("  public static String getTestValue(Integer length, Integer start, Integer len) {\n");
sb.append("    String s = \"\";\n");
sb.append("    for (int i = 0; i < length; i++) {\n");
sb.append("      s = s + \"1\";\n");
sb.append("    }\n");
sb.append("    if (start != null) {\n");
sb.append("      int startIndex = start - 1; // Java 从 0 开始\n");
sb.append("      int endIndex = startIndex + len; // 结束位置 = 起始 + 长度\n");
sb.append("      s = s.substring(startIndex, endIndex);\n");
sb.append("    }\n");
sb.append("    return s;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * sql转到json并输出到SubOut.json\n");
sb.append("   */\n");
sb.append("  public static String getJson() {\n");
sb.append("    String sql = Sub.as.getSubInTxtContent();\n");
sb.append("    String json = SqlToJsonConverter.convert(sql);\n");
sb.append("    json = convert(json);\n");
sb.append("\n");
sb.append("    Sub.u.outFile(json, Sub.as.getSubOutJsonPath());\n");
sb.append("\n");
sb.append("    return json;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 把要测试的json初始化到测试目录\n");
sb.append("   *\n");
sb.append("   * @param toTestNum 测试号\n");
sb.append("   */\n");
sb.append("  public static void jsonTableToTestInputDir(String json, String toTestNum) throws Exception {\n");
sb.append("\n");
sb.append("    // 获取测试表\n");
sb.append("    Map<String, String> map = getTestTable(json);\n");
sb.append("\n");
sb.append("    String inputPath = Sub.as.getTestDbTestCaseTestNumInputPath(toTestNum);\n");
sb.append("    String outputPath = Sub.as.getTestDbTestCaseTestNumOutputPath(toTestNum);\n");
sb.append("\n");
sb.append("    for (Map.Entry<String, String> entry : map.entrySet()) {\n");
sb.append("      String table = entry.getKey();\n");
sb.append("      // 从存储目录获取存储过的tableData.json\n");
sb.append("      String fromPath = Sub.as.getZopPath(\"testJsonCaseData\") + table + \".json\";\n");
sb.append("\n");
sb.append("      File f = new File(fromPath);\n");
sb.append("      // 如果存储目录存在此文件则放到testNum的input output目录下\n");
sb.append("      if (f.exists()) {\n");
sb.append("        String content = Sub.u.readFile(fromPath);\n");
sb.append("        Sub.u.outFileIFNotExist(content, inputPath + table + \".json\");\n");
sb.append("        Sub.u.outFileIFNotExist(content, outputPath + table + \".json\");\n");
sb.append("      } else {\n");
sb.append("        String message = \"not exist file:\" + f.getAbsolutePath();\n");
sb.append("        Sub.u.outFileAddContent(message, Sub.as.getSubOutLogPath());\n");
sb.append("        throw new Exception(message);\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 获取测试此sql用到的表\n");
sb.append("   */\n");
sb.append("  public static Map<String, String> getTestTable(String json) {\n");
sb.append("    // 1. 解析 JSON 为 Map\n");
sb.append("    Map<String, Object> root = Sub.JsonUtils.fromJson(json, new com.fasterxml.jackson.core.type.TypeReference<Map<String, Object>>() {});\n");
sb.append("    Map<String, String> map = new HashMap<String, String>();\n");
sb.append("\n");
sb.append("    // 2. 取 aliasMapping\n");
sb.append("    @SuppressWarnings(\"unchecked\")\n");
sb.append("    Map<String, String> aliasMapping = (Map<String, String>) root.get(\"aliasMapping\");\n");
sb.append("    for (Map.Entry<String, String> entry : aliasMapping.entrySet()) {\n");
sb.append("      // String alias = entry.getKey();\n");
sb.append("      String table = entry.getValue();\n");
sb.append("      map.put(table, table);\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    // 去掉重复后的table\n");
sb.append("    // for (Map.Entry<String, String> entry : map.entrySet()) {\n");
sb.append("    // String table = entry.getKey();\n");
sb.append("    // // 从存储目录获取存储过的tableData.json\n");
sb.append("    // String path = Sub.as.getZopPath(\"testJsonCaseData\") + table + \".json\";\n");
sb.append("    // File f = new File(path);\n");
sb.append("    // if (f.exists()) {\n");
sb.append("\n");
sb.append("    // } else {\n");
sb.append("    // System.out.println(\"not exist file:\" + f.getAbsolutePath());\n");
sb.append("    // }\n");
sb.append("    // }\n");
sb.append("    return map;\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 全部转换成表.字段的形式\n");
sb.append("   */\n");
sb.append("  public static String convert(String json) {\n");
sb.append("\n");
sb.append("    // 1. 解析 JSON 为 Map\n");
sb.append("    Map<String, Object> root = Sub.JsonUtils.fromJson(json, new com.fasterxml.jackson.core.type.TypeReference<Map<String, Object>>() {});\n");
sb.append("\n");
sb.append("    // 2. 取 aliasMapping\n");
sb.append("    @SuppressWarnings(\"unchecked\")\n");
sb.append("    Map<String, String> aliasMapping = (Map<String, String>) root.get(\"aliasMapping\");\n");
sb.append("\n");
sb.append("    // 3. 取 whereOrOn\n");
sb.append("    @SuppressWarnings(\"unchecked\")\n");
sb.append("    List<String> whereOrOn = (List<String>) root.get(\"whereOrOn\");\n");
sb.append("\n");
sb.append("    // 4. 转换 whereOrOn\n");
sb.append("    List<String> converted = new ArrayList<>();\n");
sb.append("\n");
sb.append("    for (String condition : whereOrOn) {\n");
sb.append("\n");
sb.append("      String result = condition;\n");
sb.append("\n");
sb.append("      for (Map.Entry<String, String> entry : aliasMapping.entrySet()) {\n");
sb.append("        String alias = entry.getKey();\n");
sb.append("        String table = entry.getValue();\n");
sb.append("\n");
sb.append("        // 只替换 t17.xxx 这种形式\n");
sb.append("        result = result.replaceAll(alias + \"\\\\.\", table + \"\\\\.\");\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      converted.add(result);\n");
sb.append("\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    // 5. 回写 whereOrOn\n");
sb.append("    root.put(\"whereOrOn\", converted);\n");
sb.append("\n");
sb.append("    return Sub.JsonUtils.toJson(root);\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 1)别名 -> 表名 映射解析器\n");
sb.append("   */\n");
sb.append("  public static class AliasMappingVisitor extends SelectVisitorAdapter {\n");
sb.append("\n");
sb.append("    private final Map<String, String> aliasMap = new LinkedHashMap<>();\n");
sb.append("\n");
sb.append("    public Map<String, String> getAliasMap() {\n");
sb.append("      return aliasMap;\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    @Override\n");
sb.append("    public void visit(PlainSelect plainSelect) {\n");
sb.append("\n");
sb.append("      // FROM t1\n");
sb.append("      if (plainSelect.getFromItem() instanceof Table table) {\n");
sb.append("        if (table.getAlias() != null) {\n");
sb.append("          aliasMap.put(table.getAlias().getName(), table.getName());\n");
sb.append("        } else {\n");
sb.append("          // 无别名 → 默认别名 = 表名（你要的效果）\n");
sb.append("          aliasMap.put(table.getName(), table.getName());\n");
sb.append("        }\n");
sb.append("\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // JOIN t2...\n");
sb.append("      if (plainSelect.getJoins() != null) {\n");
sb.append("        for (Join join : plainSelect.getJoins()) {\n");
sb.append("          if (join.getRightItem() instanceof Table joinTable) {\n");
sb.append("            if (joinTable.getAlias() != null) {\n");
sb.append("              aliasMap.put(joinTable.getAlias().getName(), joinTable.getName());\n");
sb.append("            }\n");
sb.append("          }\n");
sb.append("        }\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 2) WHERE + JOIN ON 条件解析器\n");
sb.append("   */\n");
sb.append("  public static class WhereOrOnParser extends SelectVisitorAdapter {\n");
sb.append("\n");
sb.append("    private final Map<String, String> aliasMap;\n");
sb.append("    private final List<String> result = new ArrayList<>();\n");
sb.append("\n");
sb.append("    /**\n");
sb.append("    * \n");
sb.append("    */\n");
sb.append("    public WhereOrOnParser(Map<String, String> aliasMap) {\n");
sb.append("      this.aliasMap = aliasMap;\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    public List<String> getResult() {\n");
sb.append("      return result;\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    @Override\n");
sb.append("    public void visit(PlainSelect plainSelect) {\n");
sb.append("\n");
sb.append("      if (plainSelect.getWhere() != null) {\n");
sb.append("        parseCondition(plainSelect.getWhere());\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      if (plainSelect.getJoins() != null) {\n");
sb.append("        for (Join join : plainSelect.getJoins()) {\n");
sb.append("          for (Expression on : join.getOnExpressions()) {\n");
sb.append("            parseCondition(on);\n");
sb.append("          }\n");
sb.append("        }\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    /**\n");
sb.append("     * 统一处理条件\n");
sb.append("     */\n");
sb.append("    private void parseCondition(Expression expr) {\n");
sb.append("\n");
sb.append("      // AND：拍平\n");
sb.append("      if (expr instanceof AndExpression and) {\n");
sb.append("        parseCondition(and.getLeftExpression());\n");
sb.append("        parseCondition(and.getRightExpression());\n");
sb.append("        return;\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // ✅ OR：保持结构，但内部字段转换\n");
sb.append("      if (expr instanceof OrExpression or) {\n");
sb.append("        String left = buildExpr(or.getLeftExpression());\n");
sb.append("        String right = buildExpr(or.getRightExpression());\n");
sb.append("        result.add(\"(\" + left + \") OR (\" + right + \")\");\n");
sb.append("        return;\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // =\n");
sb.append("      if (expr instanceof EqualsTo eq) {\n");
sb.append("        String left = convertExpr(eq.getLeftExpression());\n");
sb.append("        String right = convertExpr(eq.getRightExpression());\n");
sb.append("        result.add(left + \" = \" + right);\n");
sb.append("        return;\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // IN (...)\n");
sb.append("      if (expr instanceof InExpression in) {\n");
sb.append("        String left = convertExpr(in.getLeftExpression());\n");
sb.append("        ExpressionList list = (ExpressionList) in.getRightItemsList();\n");
sb.append("\n");
sb.append("        List<String> vals = new ArrayList<>();\n");
sb.append("        for (Expression e : list.getExpressions()) {\n");
sb.append("          vals.add(e.toString().replace(\"'\", \"\"));\n");
sb.append("        }\n");
sb.append("        result.add(left + \" IN (\" + String.join(\",\", vals) + \")\");\n");
sb.append("        return;\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // fallback\n");
sb.append("      result.add(expr.toString());\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    /**\n");
sb.append("    * \n");
sb.append("    */\n");
sb.append("    private String buildExpr(Expression expr) {\n");
sb.append("\n");
sb.append("      if (expr instanceof AndExpression and) {\n");
sb.append("        return buildExpr(and.getLeftExpression()) + \"\\n  AND \" + buildExpr(and.getRightExpression());\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      if (expr instanceof EqualsTo eq) {\n");
sb.append("        return convertExpr(eq.getLeftExpression()) + \" = \" + convertExpr(eq.getRightExpression());\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      if (expr instanceof InExpression in) {\n");
sb.append("        String left = convertExpr(in.getLeftExpression());\n");
sb.append("        ExpressionList list = (ExpressionList) in.getRightItemsList();\n");
sb.append("\n");
sb.append("        List<String> vals = new ArrayList<>();\n");
sb.append("        for (Expression e : list.getExpressions()) {\n");
sb.append("          vals.add(e.toString().replace(\"'\", \"\"));\n");
sb.append("        }\n");
sb.append("        return left + \" IN (\" + String.join(\",\", vals) + \")\";\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // 兜底\n");
sb.append("      return expr.toString();\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    /**\n");
sb.append("     * 字段/函数转换器 (含 SUBSTR)\n");
sb.append("     */\n");
sb.append("    private String convertExpr(Expression expr) {\n");
sb.append("\n");
sb.append("      // SUBSTR(t1.xxx,1,5)\n");
sb.append("      if (expr instanceof Function func && func.getName().equalsIgnoreCase(\"SUBSTR\")) {\n");
sb.append("        List<Expression> args = func.getParameters().getExpressions();\n");
sb.append("\n");
sb.append("        String col = args.get(0).toString();\n");
sb.append("        String start = args.get(1).toString();\n");
sb.append("        String len = args.get(2).toString();\n");
sb.append("\n");
sb.append("        String[] arr = col.split(\"\\\\.\");\n");
sb.append("        String alias = arr[0];\n");
sb.append("        String field = arr[1];\n");
sb.append("\n");
sb.append("        String table = aliasMap.getOrDefault(alias, alias);\n");
sb.append("\n");
sb.append("        return table + \".\" + field + \"#\" + start + \",\" + len + \"#\";\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      // t1.xxx → TB_XXX.xxx\n");
sb.append("      String text = expr.toString();\n");
sb.append("      if (text.contains(\".\")) {\n");
sb.append("        String[] arr = text.split(\"\\\\.\");\n");
sb.append("        String alias = arr[0];\n");
sb.append("        String field = arr[1];\n");
sb.append("\n");
sb.append("        String table = aliasMap.getOrDefault(alias, alias);\n");
sb.append("        return table + \".\" + field;\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      return text;\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 最终 SQL -> JSON 转换器\n");
sb.append("   */\n");
sb.append("  public static class SqlToJsonConverter {\n");
sb.append("\n");
sb.append("    /**\n");
sb.append("    * \n");
sb.append("    */\n");
sb.append("    public static String convert(String sql) {\n");
sb.append("\n");
sb.append("      // 兼容mybaits\n");
sb.append("      sql = sql.replaceAll(\"#\\\\{[^}]+}\", \":param\");\n");
sb.append("\n");
sb.append("      Statement stmt = null;\n");
sb.append("      try {\n");
sb.append("        stmt = CCJSqlParserUtil.parse(sql);\n");
sb.append("      } catch (JSQLParserException e) {\n");
sb.append("        e.printStackTrace();\n");
sb.append("      }\n");
sb.append("      Select select = (Select) stmt;\n");
sb.append("\n");
sb.append("      Map<String, Object> json = new LinkedHashMap<>();\n");
sb.append("\n");
sb.append("      // 1. 别名 mapping\n");
sb.append("      AliasMappingVisitor aliasVisitor = new AliasMappingVisitor();\n");
sb.append("      select.getSelectBody().accept(aliasVisitor);\n");
sb.append("      Map<String, String> aliasMap = aliasVisitor.getAliasMap();\n");
sb.append("      json.put(\"aliasMapping\", aliasMap);\n");
sb.append("\n");
sb.append("      // 2. WHERE + JOIN 条件\n");
sb.append("      WhereOrOnParser wparser = new WhereOrOnParser(aliasMap);\n");
sb.append("      select.getSelectBody().accept(wparser);\n");
sb.append("\n");
sb.append("      json.put(\"whereOrOn\", wparser.getResult());\n");
sb.append("\n");
sb.append("      return Sub.JsonUtils.toJson(json);\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("}\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
return sb.toString();
    }
    /**
     * 
     */
    public String getStartXUTester() {
StringBuilder sb=new StringBuilder();
sb.append("////////////////////////////////////////////////////////////////////////////\n");
sb.append("//\n");
sb.append("// プログラム名：仮口座削除\n");
sb.append("// ファイル名：SBMAB002\n");
sb.append("// 作成者：徐楠\n");
sb.append("// 作成日：2025/12/16\n");
sb.append("// 備考：なし\n");
sb.append("// 変更履歴：なし\n");
sb.append("// 2025/12/16 徐楠 新規作成\n");
sb.append("//\n");
sb.append("////////////////////////////////////////////////////////////////////////////\n");
sb.append("\n");
sb.append("package " + Sub.class.getPackage().getName()+  ".sub;\n");
sb.append("\n");
sb.append("\n");
sb.append("import jp.co.skywave.ut.UTBase;\n");
sb.append("import jp.co.skywave.ut.scenario.UTConcierge;\n");
sb.append("import jp.co.skywave.ut.scenario.UTContext;\n");
sb.append("import jp.or.jasdec.sbf.ap.parts.exception.APZZPartsRetryException;\n");
sb.append("import jp.or.jasdec.sbf.ap.parts.file.APZZFileOperation;\n");
sb.append("import jp.or.jasdec.sbf.ap.parts.log.APZZLogOutStub;\n");
sb.append("import jp.or.jasdec.sbf.ap.parts.springannotation.APZZAutoWired;\n");
sb.append("import jp.or.jasdec.sbf.sb.ma.batch.biz.Sub;\n");
sb.append("import jp.or.jasdec.sbf.sb.ma.parts.bean.db.SBMAQSelectTraCancel002InDataBean;\n");
sb.append("import jp.or.jasdec.sbf.sb.ma.parts.bean.db.SBMAQUpdateReqMashoInfo002InDataBean;\n");
sb.append("import jp.or.jasdec.sbf.sb.ma.parts.mapper.SBMAMapper;\n");
sb.append("import java.sql.DriverManager;\n");
sb.append("import java.sql.SQLException;\n");
sb.append("import java.sql.Statement;\n");
sb.append("import java.util.Map;\n");
sb.append("import org.junit.jupiter.api.Tag;\n");
sb.append("import org.junit.jupiter.api.Test;\n");
sb.append("import org.junit.jupiter.api.TestInstance;\n");
sb.append("import org.mockito.ArgumentCaptor;\n");
sb.append("import org.mockito.Captor;\n");
sb.append("import org.slf4j.Logger;\n");
sb.append("import org.slf4j.LoggerFactory;\n");
sb.append("import org.springframework.beans.factory.annotation.Autowired;\n");
sb.append("import org.springframework.boot.test.context.SpringBootTest;\n");
sb.append("import org.springframework.core.env.Environment;\n");
sb.append("import org.springframework.test.context.bean.override.mockito.MockitoSpyBean;\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("/**\n");
sb.append(" * サブシステム[SBMA:SB管理]の業務機能[SBMAB002:仮口座削除] のテスト実行クラスです。\n");
sb.append(" */\n");
sb.append("@SpringBootTest\n");
sb.append("@Tag(\"subSysId:SBMA\")\n");
sb.append("@Tag(\"businessFunctionId:SBMAB002\")\n");
sb.append("@Tag(\"fwType:バッチ\")\n");
sb.append("@Tag(\"bizClassBeanId:jp.or.jasdec.sbf.sb.ma.batch.biz.SBMAB002\")\n");
sb.append("@Tag(\"bizInitClassBeanId:jp.or.jasdec.sbf.sb.zz.parts.biz.base.SBZZInitialStartProcess\")\n");
sb.append("@TestInstance(TestInstance.Lifecycle.PER_CLASS)\n");
sb.append("public class StartXUTester extends UTBase {\n");
sb.append("  private static final Logger logger = LoggerFactory.getLogger(StartXUTester.class);\n");
sb.append("  /**\n");
sb.append("   * テストコンシェルジュクラスインスタンス\n");
sb.append("   */\n");
sb.append("  @Autowired\n");
sb.append("  private UTConcierge concierge;\n");
sb.append("  @MockitoSpyBean\n");
sb.append("  private SBMAMapper sbmaMap;\n");
sb.append("  @APZZAutoWired\n");
sb.append("  private APZZLogOutStub logOutStub;\n");
sb.append("  @MockitoSpyBean\n");
sb.append("  private APZZFileOperation apzzFileOperation;\n");
sb.append("  @Captor\n");
sb.append("  private ArgumentCaptor<SBMAQSelectTraCancel002InDataBean> inCaptor;\n");
sb.append("  @Captor\n");
sb.append("  private ArgumentCaptor<SBMAQUpdateReqMashoInfo002InDataBean> captor;\n");
sb.append("  @Autowired\n");
sb.append("  private Environment env;\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * テスト全体前処理です。\n");
sb.append("   */\n");
sb.append("  @Override\n");
sb.append("  protected void beforeAllImpl(UTContext context) throws Exception {\n");
sb.append("    Sub.as.clearLog();\n");
sb.append("    clear();\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * テスト全体後処理です。\n");
sb.append("   */\n");
sb.append("  @Override\n");
sb.append("  protected void afterAllImpl(UTContext context) throws Exception {\n");
sb.append("    // 処理なし\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * テストケース別前処理です。\n");
sb.append("   */\n");
sb.append("  @Override\n");
sb.append("  protected void beforeEachImpl(UTContext context) throws Exception {\n");
sb.append("    // 処理なし\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * テストケース別後処理です。\n");
sb.append("   */\n");
sb.append("  @Override\n");
sb.append("  protected void afterEachImpl(UTContext context) throws Exception {\n");
sb.append("    // 処理なし\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * データをクリアする\n");
sb.append("   */\n");
sb.append("  private void clear() {\n");
sb.append("    // 処理なし\n");
sb.append("    String[] sqls = {\"TB_FSBREQHIST\"};\n");
sb.append("    Statement stmt;\n");
sb.append("    try {\n");
sb.append("      stmt = DriverManager.getConnection(env.getProperty(\"spring.datasource.url\"), env.getProperty(\"spring.datasource.username\"),\n");
sb.append("          env.getProperty(\"spring.datasource.password\")).createStatement();\n");
sb.append("      for (String s : sqls) {\n");
sb.append("        stmt.addBatch(\" delete from \" + s);\n");
sb.append("      }\n");
sb.append("      stmt.executeBatch();\n");
sb.append("    } catch (SQLException e) {\n");
sb.append("      e.printStackTrace();\n");
sb.append("    }\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テストケース[111111]を実行します。\n");
sb.append("   *\n");
sb.append("   * <pre>\n");
sb.append("   * 振替申請取消情報一括取得\n");
sb.append("   * gradle test --tests \"StartXUTester.test111111\"\n");
sb.append("   * </pre>\n");
sb.append("   *\n");
sb.append("   */\n");
sb.append("  @Test\n");
sb.append("  @Tag(\"testCaseId:111111\")\n");
sb.append("  public void test111111() throws APZZPartsRetryException {\n");
sb.append("    this.concierge.setResultNormal(true);\n");
sb.append("    this.concierge.setDuplicateCheckFlag(true);\n");
sb.append("    logger.info(\"テストケース[111111]実行開始\");\n");
sb.append("    this.concierge.invokeTestCase();\n");
sb.append("    logger.info(\"テストケース[111111]実行終了\");\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テストケース[111112]を実行します。\n");
sb.append("   *\n");
sb.append("   * <pre>\n");
sb.append("   * 買入消却申請取消\n");
sb.append("   * gradle test --tests \"StartXUTester.test111112\"\n");
sb.append("   * </pre>\n");
sb.append("   *\n");
sb.append("   */\n");
sb.append("  @Test\n");
sb.append("  @Tag(\"testCaseId:111112\")\n");
sb.append("  public void test111112() throws APZZPartsRetryException {\n");
sb.append("    this.concierge.setResultNormal(true);\n");
sb.append("    this.concierge.setDuplicateCheckFlag(true);\n");
sb.append("    logger.info(\"テストケース[111112]実行開始\");\n");
sb.append("    this.concierge.invokeTestCase();\n");
sb.append("    logger.info(\"テストケース[111112]実行終了\");\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テストケース[111113]を実行します。\n");
sb.append("   *\n");
sb.append("   * <pre>\n");
sb.append("   * 新規記録申請取消\n");
sb.append("   * gradle test --tests \"StartXUTester.test111113\"\n");
sb.append("   * </pre>\n");
sb.append("   *\n");
sb.append("   */\n");
sb.append("  @Test\n");
sb.append("  @Tag(\"testCaseId:111113\")\n");
sb.append("  public void test111113() throws APZZPartsRetryException {\n");
sb.append("    this.concierge.setResultNormal(true);\n");
sb.append("    this.concierge.setDuplicateCheckFlag(true);\n");
sb.append("    logger.info(\"テストケース[111113]実行開始\");\n");
sb.append("    this.concierge.invokeTestCase();\n");
sb.append("    logger.info(\"テストケース[111113]実行終了\");\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テストケース[111114]を実行します。\n");
sb.append("   *\n");
sb.append("   * <pre>\n");
sb.append("   * 抹消申請情報取消\n");
sb.append("   * gradle test --tests \"StartXUTester.111114\"\n");
sb.append("   * </pre>\n");
sb.append("   *\n");
sb.append("   */\n");
sb.append("  @Test\n");
sb.append("  @Tag(\"testCaseId:111114\")\n");
sb.append("  public void test111114() throws APZZPartsRetryException {\n");
sb.append("    this.concierge.setResultNormal(true);\n");
sb.append("    this.concierge.setDuplicateCheckFlag(true);\n");
sb.append("    logger.info(\"テストケース[111114]実行開始\");\n");
sb.append("    this.concierge.invokeTestCase();\n");
sb.append("    logger.info(\"テストケース[111114]実行終了\");\n");
sb.append("  }\n");
sb.append("\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   *\n");
sb.append("   * テストケース[111115]を実行します。\n");
sb.append("   *\n");
sb.append("   * <pre>\n");
sb.append("   * 日銀GW送信無し取得\n");
sb.append("   * gradle test --tests \"StartXUTester.111115\"\n");
sb.append("   * </pre>\n");
sb.append("   *\n");
sb.append("   */\n");
sb.append("  @Test\n");
sb.append("  @Tag(\"testCaseId:111115\")\n");
sb.append("  public void test111115() throws APZZPartsRetryException {\n");
sb.append("    this.concierge.setResultNormal(true);\n");
sb.append("    this.concierge.setDuplicateCheckFlag(true);\n");
sb.append("    logger.info(\"テストケース[111115]実行開始\");\n");
sb.append("    this.concierge.invokeTestCase();\n");
sb.append("    logger.info(\"テストケース[111115]実行終了\");\n");
sb.append("  }\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("}\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
return sb.toString();
    }
    /**
     * 
     */
    public String getIndex() {
StringBuilder sb=new StringBuilder();
sb.append("package " + Sub.class.getPackage().getName()+  ".sub;\n");
sb.append("\n");
sb.append("import jp.or.jasdec.sbf.sb.ma.batch.biz.Sub;\n");
sb.append("\n");
sb.append("public class Index {\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * VsCode setting.json\n");
sb.append("   */\n");
sb.append("  public static String settings() {\n");
sb.append("\n");
sb.append("    return \"\"\"\n");
sb.append("        //缩进tab换成空格\n");
sb.append("        \"[java]\": {\n");
sb.append("        \"editor.tabSize\": 2,\n");
sb.append("        \"editor.insertSpaces\": true,\n");
sb.append("        \"editor.detectIndentation\": false\n");
sb.append("        },\n");
sb.append("\n");
sb.append("          \"\"\";\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  public void Util() {\n");
sb.append("\n");
sb.append("    // 获取工程名 SBMAB002\n");
sb.append("    Sub.u.getProjectName();\n");
sb.append("\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  public void ApplicationSub() {\n");
sb.append("\n");
sb.append("    // 電文の期待値ファイルのパス\n");
sb.append("    String caseNum = \"111111\";\n");
sb.append("    // C:\\opt\\workspace\\SBMAB002\\test\\inputData\\SBMA\\SBMAB002\\telegram\\testCase\\111111\\expect\\expect.json\n");
sb.append("    String expectFilePath = Sub.as.getJsonTelegramExpectPath(caseNum);\n");
sb.append("\n");
sb.append("    // 获取电文路径 C:\\opt\\workspace\\SBMAB002\\test\\inputData\\SBMA\\SBMAB002\\telegram\\data\\\\\n");
sb.append("    String path = Sub.as.getTelegramDataPath();\n");
sb.append("\n");
sb.append("    // 获取期待之中的电文索引列表 list ← SBJMO049_1071\n");
sb.append("    String json = \"{\\\"expectTestDataIDs\\\": [ \\\"SBJMO049_1071\\\"]}\";\n");
sb.append("    Sub.as.getTelegramExpectList(json);\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * SubのContent\n");
sb.append("   */\n");
sb.append("  public void content() {\n");
sb.append("    // 期待値の初期化の内容\n");
sb.append("    String json = Sub.c.telegramExpectJson();\n");
sb.append("  }\n");
sb.append("}\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
return sb.toString();
    }
    /**
     * 
     */
    public String getUTDbUtil() {
StringBuilder sb=new StringBuilder();
sb.append("// // 打印具体不对的地方\n");
sb.append("// // Source code is decompiled from a .class file using FernFlower decompiler (from Intellij IDEA).\n");
sb.append("package " + Sub.class.getPackage().getName()+  ".sub;\n");
sb.append("package " + Sub.class.getPackage().getName()+  ".sub;\n");
sb.append("\n");
sb.append("// import com.fasterxml.jackson.core.JsonProcessingException;\n");
sb.append("// import com.fasterxml.jackson.databind.DeserializationFeature;\n");
sb.append("// import com.fasterxml.jackson.databind.ObjectMapper;\n");
sb.append("// import java.io.BufferedReader;\n");
sb.append("// import java.io.File;\n");
sb.append("// import java.io.FileReader;\n");
sb.append("// import java.io.IOException;\n");
sb.append("// import java.math.BigDecimal;\n");
sb.append("// import java.sql.SQLException;\n");
sb.append("// import java.time.LocalDate;\n");
sb.append("// import java.time.LocalDateTime;\n");
sb.append("// import java.time.format.DateTimeFormatter;\n");
sb.append("// import java.time.format.DateTimeParseException;\n");
sb.append("// import java.util.ArrayList;\n");
sb.append("// import java.util.Arrays;\n");
sb.append("// import java.util.HashMap;\n");
sb.append("// import java.util.HashSet;\n");
sb.append("// import java.util.Iterator;\n");
sb.append("// import java.util.LinkedHashMap;\n");
sb.append("// import java.util.List;\n");
sb.append("// import java.util.Map;\n");
sb.append("// import java.util.Objects;\n");
sb.append("// import java.util.Random;\n");
sb.append("// import java.util.Set;\n");
sb.append("// import java.util.stream.Collectors;\n");
sb.append("// import java.util.stream.Stream;\n");
sb.append("// import jp.co.skywave.ut.db.bean.ColumnDefinitionBean;\n");
sb.append("// import jp.co.skywave.ut.db.bean.TableDataSetBean;\n");
sb.append("// import jp.co.skywave.ut.db.bean.TableDefinitionBean;\n");
sb.append("// import jp.co.skywave.ut.scenario.UTContext;\n");
sb.append("// import jp.co.skywave.ut.util.UTFolderUtil;\n");
sb.append("// import jp.or.jasdec.sbf.sb.ma.batch.biz.Sub.JsonUtils;\n");
sb.append("// import org.dbunit.DatabaseUnitException;\n");
sb.append("// import org.dbunit.assertion.DbUnitAssert;\n");
sb.append("// import org.dbunit.assertion.DiffCollectingFailureHandler;\n");
sb.append("// import org.dbunit.assertion.Difference;\n");
sb.append("// import org.dbunit.database.IDatabaseConnection;\n");
sb.append("// import org.dbunit.dataset.Column;\n");
sb.append("// import org.dbunit.dataset.DataSetException;\n");
sb.append("// import org.dbunit.dataset.DefaultDataSet;\n");
sb.append("// import org.dbunit.dataset.DefaultTable;\n");
sb.append("// import org.dbunit.dataset.IDataSet;\n");
sb.append("// import org.dbunit.dataset.ITable;\n");
sb.append("// import org.dbunit.dataset.ITableMetaData;\n");
sb.append("// import org.dbunit.dataset.SortedTable;\n");
sb.append("// import org.dbunit.dataset.datatype.DataType;\n");
sb.append("// import org.slf4j.Logger;\n");
sb.append("// import org.slf4j.LoggerFactory;\n");
sb.append("\n");
sb.append("// public class UTDbUtil {\n");
sb.append("// private static final String IGNORE_TEST_CASE_DB_ID_FILE = \"./test/ignore_testcase_db_id.txt\";\n");
sb.append("// private static final Logger logger = LoggerFactory.getLogger(UTDbUtil.class);\n");
sb.append("// private static final Random RANDOM = new Random();\n");
sb.append("// private static final DateTimeFormatter DATE_FORMATTER_FOR_CHECK_YYYYMD =\n");
sb.append("// DateTimeFormatter.ofPattern(\"yyyy/M/d\");\n");
sb.append("// private static final DateTimeFormatter DATE_FORMATTER_FOR_CHECK_YYYYMMDD =\n");
sb.append("// DateTimeFormatter.ofPattern(\"yyyy/MM/dd\");\n");
sb.append("// private static final DateTimeFormatter DATE_FORMATTER_FOR_BUILD =\n");
sb.append("// DateTimeFormatter.ofPattern(\"yyyy-MM-dd\");\n");
sb.append("// private static final DateTimeFormatter TIMESTAMP_FORMATTER_CHECK_YYYYMDHMS =\n");
sb.append("// DateTimeFormatter.ofPattern(\"yyyy/M/d H:m:s\");\n");
sb.append("// private static final DateTimeFormatter TIMESTAMP_FORMATTER_CHECK_YYYYMMDDHHMMSS =\n");
sb.append("// DateTimeFormatter.ofPattern(\"yyyy/MM/dd HH:mm:ss\");\n");
sb.append("// private static final DateTimeFormatter TIMESTAMP_FORMATTER_BUILD =\n");
sb.append("// DateTimeFormatter.ofPattern(\"yyyy-MM-dd HH:mm:ss\");\n");
sb.append("\n");
sb.append("// public UTDbUtil() {}\n");
sb.append("\n");
sb.append("// public static void loadDbDefineSetting(UTContext context) {\n");
sb.append("// context.setTableDefineMap(new HashMap());\n");
sb.append("// String dbDefineFolder = UTFolderUtil.getInputDbDefineFolder(context);\n");
sb.append("// File dbDefileFolderFile = new File(dbDefineFolder);\n");
sb.append("// if (dbDefileFolderFile.exists()) {\n");
sb.append("// File[] dbDefineFiles = dbDefileFolderFile.listFiles();\n");
sb.append("// File[] var4 = dbDefineFiles;\n");
sb.append("// int var5 = dbDefineFiles.length;\n");
sb.append("\n");
sb.append("// for (int var6 = 0; var6 < var5; ++var6) {\n");
sb.append("// File dbDefineFile = var4[var6];\n");
sb.append("// TableDefinitionBean tableDefinitionBean = convertTableDefineJsonToBean(dbDefineFile.getPath());\n");
sb.append("// String physicalTableName = tableDefinitionBean.getPhysicalTableName();\n");
sb.append("// context.getTableDefineMap().put(physicalTableName, tableDefinitionBean);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static void loadDbMasterDataSetting(UTContext context) {\n");
sb.append("// context.setMasterTableDataMap(new HashMap());\n");
sb.append("// String masterDbFolder = UTFolderUtil.getInputDbMasterFolder(context);\n");
sb.append("// File masterDbFolderFile = new File(masterDbFolder);\n");
sb.append("// if (masterDbFolderFile.exists()) {\n");
sb.append("// File[] masterDbFiles = masterDbFolderFile.listFiles();\n");
sb.append("// File[] var4 = masterDbFiles;\n");
sb.append("// int var5 = masterDbFiles.length;\n");
sb.append("\n");
sb.append("// for (int var6 = 0; var6 < var5; ++var6) {\n");
sb.append("// File masterDbFile = var4[var6];\n");
sb.append("// TableDataSetBean masterTableDataSetBean = convertTableDataSet(masterDbFile.getPath());\n");
sb.append("// String physicalTableName = masterTableDataSetBean.getPhysicalTableName();\n");
sb.append("// context.getMasterTableDataMap().put(physicalTableName, masterTableDataSetBean);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static void loadDbTestCaseDataSetting(UTContext context) {\n");
sb.append("// context.setTestCasePrepareTableDataMap(new HashMap());\n");
sb.append("// String testCaseDbFolder = UTFolderUtil.getInputDbTestCasePrepareFolder(context);\n");
sb.append("// File testCaseDbFolderFile = new File(testCaseDbFolder);\n");
sb.append("// File[] testCaseDbFiles;\n");
sb.append("// File[] var4;\n");
sb.append("// int var5;\n");
sb.append("// int var6;\n");
sb.append("// File testCaseDbFile;\n");
sb.append("// TableDataSetBean tableDataSetBean;\n");
sb.append("// String physicalTableName;\n");
sb.append("// if (testCaseDbFolderFile.exists()) {\n");
sb.append("// testCaseDbFiles = testCaseDbFolderFile.listFiles();\n");
sb.append("// var4 = testCaseDbFiles;\n");
sb.append("// var5 = testCaseDbFiles.length;\n");
sb.append("\n");
sb.append("// for (var6 = 0; var6 < var5; ++var6) {\n");
sb.append("// testCaseDbFile = var4[var6];\n");
sb.append("// tableDataSetBean = convertTableDataSet(testCaseDbFile.getPath());\n");
sb.append("// physicalTableName = tableDataSetBean.getPhysicalTableName();\n");
sb.append("// context.getTestCasePrepareTableDataMap().put(physicalTableName, tableDataSetBean);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// context.setTestCaseExpectTableDataMap(new HashMap());\n");
sb.append("// testCaseDbFolder = UTFolderUtil.getInputDbTestCaseExpectFolder(context);\n");
sb.append("// testCaseDbFolderFile = new File(testCaseDbFolder);\n");
sb.append("// if (testCaseDbFolderFile.exists()) {\n");
sb.append("// testCaseDbFiles = testCaseDbFolderFile.listFiles();\n");
sb.append("// var4 = testCaseDbFiles;\n");
sb.append("// var5 = testCaseDbFiles.length;\n");
sb.append("\n");
sb.append("// for (var6 = 0; var6 < var5; ++var6) {\n");
sb.append("// testCaseDbFile = var4[var6];\n");
sb.append("// tableDataSetBean = convertTableDataSet(testCaseDbFile.getPath());\n");
sb.append("// physicalTableName = tableDataSetBean.getPhysicalTableName();\n");
sb.append("// context.getTestCaseExpectTableDataMap().put(physicalTableName, tableDataSetBean);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static TableDefinitionBean convertTableDefineJsonToBean(String tableDefineFileName) {\n");
sb.append("// try {\n");
sb.append("// ObjectMapper mapper = new ObjectMapper();\n");
sb.append("// mapper.enable(DeserializationFeature.USE_BIG_DECIMAL_FOR_FLOATS);\n");
sb.append("// TableDefinitionBean tableDefinition = (TableDefinitionBean) mapper.readValue(new\n");
sb.append("// File(tableDefineFileName), TableDefinitionBean.class);\n");
sb.append("// return tableDefinition;\n");
sb.append("// } catch (IOException var3) {\n");
sb.append("// String message = String.format(\"テーブル定義ファイル[%s]のJSON形式からBean形式への変換に失敗しました。\", tableDefineFileName);\n");
sb.append("// throw new RuntimeException(message, var3);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static TableDataSetBean convertTableDataSet(String tableDataFileName) {\n");
sb.append("// File tableDataFileObj = new File(tableDataFileName);\n");
sb.append("// if (tableDataFileObj.exists() && tableDataFileObj.length() != 0L) {\n");
sb.append("// try {\n");
sb.append("// ObjectMapper mapper = new ObjectMapper();\n");
sb.append("// mapper.enable(DeserializationFeature.USE_BIG_DECIMAL_FOR_FLOATS);\n");
sb.append("// TableDataSetBean dataSet = (TableDataSetBean) mapper.readValue(new File(tableDataFileName),\n");
sb.append("// TableDataSetBean.class);\n");
sb.append("// return dataSet;\n");
sb.append("// } catch (IOException var4) {\n");
sb.append("// String message = String.format(\"テーブルデータファイル[%s]のJSON形式からBean形式への変換に失敗しました。\", tableDataFileName);\n");
sb.append("// throw new RuntimeException(message, var4);\n");
sb.append("// }\n");
sb.append("// } else {\n");
sb.append("// TableDataSetBean dataSet = new TableDataSetBean();\n");
sb.append("// dataSet.setPhysicalTableName(getBaseName(tableDataFileObj.getName()));\n");
sb.append("// dataSet.setTableData(new ArrayList());\n");
sb.append("// return dataSet;\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static IDataSet buildMasterPrepareTableData(UTContext context) {\n");
sb.append("// Map<String, TableDefinitionBean> tableDefineMap = context.getTableDefineMap();\n");
sb.append("// Map<String, TableDataSetBean> masterTableDataMap = context.getMasterTableDataMap();\n");
sb.append("// if (masterTableDataMap.size() == 0) {\n");
sb.append("// return null;\n");
sb.append("// } else {\n");
sb.append("// DefaultDataSet dataSet = new DefaultDataSet();\n");
sb.append("// Iterator var4 = masterTableDataMap.entrySet().iterator();\n");
sb.append("\n");
sb.append("// while (var4.hasNext()) {\n");
sb.append("// Map.Entry<String, TableDataSetBean> entry = (Map.Entry) var4.next();\n");
sb.append("// TableDefinitionBean tableDefinition = (TableDefinitionBean) tableDefineMap.get(entry.getKey());\n");
sb.append("// TableDataSetBean tableDataSet = (TableDataSetBean) entry.getValue();\n");
sb.append("// ITable table = buildTableData(context, tableDefinition, tableDataSet);\n");
sb.append("\n");
sb.append("// try {\n");
sb.append("// logger.info(\n");
sb.append("// String.format(\"マスタ対象テーブル[%s:%s]のテーブルセット作成を開始します。\", tableDefinition.getPhysicalTableName(),\n");
sb.append("// tableDefinition.getLogicalTableName()));\n");
sb.append("// dataSet.addTable(table);\n");
sb.append("// logger.info(\n");
sb.append("// String.format(\"マスタ対象テーブル[%s:%s]のテーブルセット作成が完了しました。\", tableDefinition.getPhysicalTableName(),\n");
sb.append("// tableDefinition.getLogicalTableName()));\n");
sb.append("// } catch (DataSetException var11) {\n");
sb.append("// String message = String.format(\"マスタ対象テーブル[%s]のテーブルセット作成に失敗しました。\",\n");
sb.append("// tableDefinition.getPhysicalTableName());\n");
sb.append("// throw new RuntimeException(message, var11);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return dataSet;\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static IDataSet buildTestCasePrepareTableData(UTContext context) {\n");
sb.append("// Map<String, TableDefinitionBean> tableDefineMap = context.getTableDefineMap();\n");
sb.append("// Map<String, TableDataSetBean> tableDataMap = context.getTestCasePrepareTableDataMap();\n");
sb.append("// if (tableDataMap.size() == 0) {\n");
sb.append("// return null;\n");
sb.append("// } else {\n");
sb.append("// Map<String, Set<String>> ignoreTestCaseTableMap = (Map)\n");
sb.append("// context.getIgnoreTestCaseTableMap().get(context.getTestCaseId());\n");
sb.append("// if (ignoreTestCaseTableMap == null) {\n");
sb.append("// ignoreTestCaseTableMap = new HashMap();\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// DefaultDataSet dataSet = new DefaultDataSet();\n");
sb.append("// Iterator var5 = tableDataMap.entrySet().iterator();\n");
sb.append("\n");
sb.append("// while (true) {\n");
sb.append("// while (var5.hasNext()) {\n");
sb.append("// Map.Entry<String, TableDataSetBean> entry = (Map.Entry) var5.next();\n");
sb.append("// logger.info(\"&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&==============\" + entry.getKey());\n");
sb.append("// TableDefinitionBean tableDefinition = (TableDefinitionBean) tableDefineMap.get(entry.getKey());\n");
sb.append("// Set<String> opeSet = (Set) ((Map)\n");
sb.append("// ignoreTestCaseTableMap).get(tableDefinition.getPhysicalTableName());\n");
sb.append("// if (opeSet == null) {\n");
sb.append("// opeSet = new HashSet();\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// if (!((Set) opeSet).contains(\"all\") && !((Set) opeSet).contains(\"input\")) {\n");
sb.append("// TableDataSetBean tableDataSet = (TableDataSetBean) entry.getValue();\n");
sb.append("// ITable table = buildTableData(context, tableDefinition, tableDataSet);\n");
sb.append("\n");
sb.append("// try {\n");
sb.append("// logger.info(String.format(\"テストケース[%s]入力対象テーブル[%s:%s]のテーブルセット作成を開始します。\", context.getTestCaseId(),\n");
sb.append("// tableDefinition.getPhysicalTableName(),\n");
sb.append("// tableDefinition.getLogicalTableName()));\n");
sb.append("// dataSet.addTable(table);\n");
sb.append("// logger.info(String.format(\"テストケース[%s]入力対象テーブル[%s:%s]のテーブルセット作成が完了しました。\", context.getTestCaseId(),\n");
sb.append("// tableDefinition.getPhysicalTableName(), tableDefinition.getLogicalTableName()));\n");
sb.append("// } catch (DataSetException var13) {\n");
sb.append("// String message = String.format(\"テストケース準備対象テーブル[%s]のテーブルセット作成に失敗しました。\",\n");
sb.append("// tableDefinition.getPhysicalTableName());\n");
sb.append("// throw new RuntimeException(message, var13);\n");
sb.append("// }\n");
sb.append("// } else {\n");
sb.append("// logger.info(String.format(\"テストケース[%s]入力対象テーブル[%s:%s]は無視対象のためスキップします。\", context.getTestCaseId(),\n");
sb.append("// tableDefinition.getPhysicalTableName(),\n");
sb.append("// tableDefinition.getLogicalTableName()));\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return dataSet;\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static IDataSet buildTestCaseExpectTableData(UTContext context) {\n");
sb.append("// Map<String, TableDefinitionBean> tableDefineMap = context.getTableDefineMap();\n");
sb.append("// Map<String, TableDataSetBean> tableDataMap = context.getTestCaseExpectTableDataMap();\n");
sb.append("// if (tableDataMap.size() == 0) {\n");
sb.append("// return null;\n");
sb.append("// } else {\n");
sb.append("// Map<String, Set<String>> ignoreTestCaseTableMap = (Map)\n");
sb.append("// context.getIgnoreTestCaseTableMap().get(context.getTestCaseId());\n");
sb.append("// if (ignoreTestCaseTableMap == null) {\n");
sb.append("// ignoreTestCaseTableMap = new HashMap();\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// DefaultDataSet dataSet = new DefaultDataSet();\n");
sb.append("// Iterator var5 = tableDataMap.entrySet().iterator();\n");
sb.append("\n");
sb.append("// while (true) {\n");
sb.append("// while (var5.hasNext()) {\n");
sb.append("// Map.Entry<String, TableDataSetBean> entry = (Map.Entry) var5.next();\n");
sb.append("// logger.info(\"@@@@@@@@@@@@@@@@@@@@@@==============\" + entry.getKey());\n");
sb.append("// TableDefinitionBean tableDefinition = (TableDefinitionBean) tableDefineMap.get(entry.getKey());\n");
sb.append("// Set<String> opeSet = (Set) ((Map)\n");
sb.append("// ignoreTestCaseTableMap).get(tableDefinition.getPhysicalTableName());\n");
sb.append("// if (opeSet == null) {\n");
sb.append("// opeSet = new HashSet();\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// if (!((Set) opeSet).contains(\"all\") && !((Set) opeSet).contains(\"expect\")) {\n");
sb.append("// TableDataSetBean tableDataSet = (TableDataSetBean) entry.getValue();\n");
sb.append("// ITable table = buildTableData(context, tableDefinition, tableDataSet);\n");
sb.append("\n");
sb.append("// try {\n");
sb.append("// dataSet.addTable(table);\n");
sb.append("// } catch (DataSetException var13) {\n");
sb.append("// String message = String.format(\"テストケース期待対象テーブル[%s]のテーブルセット作成に失敗しました。\",\n");
sb.append("// tableDefinition.getPhysicalTableName());\n");
sb.append("// throw new RuntimeException(message, var13);\n");
sb.append("// }\n");
sb.append("// } else {\n");
sb.append("// logger.info(String.format(\"テストケース[%s]期待対象テーブル[%s:%s]は無視対象のためスキップします。\", context.getTestCaseId(),\n");
sb.append("// tableDefinition.getPhysicalTableName(),\n");
sb.append("// tableDefinition.getLogicalTableName()));\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return dataSet;\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static IDataSet buildTestCaseActualTableData(UTContext context, IDatabaseConnection\n");
sb.append("// connection) {\n");
sb.append("// Map<String, TableDataSetBean> tableDataMap = context.getTestCaseExpectTableDataMap();\n");
sb.append("// if (tableDataMap.size() == 0) {\n");
sb.append("// return null;\n");
sb.append("// } else {\n");
sb.append("// List<String> tables = new ArrayList();\n");
sb.append("// Iterator var4 = tableDataMap.entrySet().iterator();\n");
sb.append("\n");
sb.append("// while (var4.hasNext()) {\n");
sb.append("// Map.Entry<String, TableDataSetBean> entry = (Map.Entry) var4.next();\n");
sb.append("// tables.add((String) entry.getKey());\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// try {\n");
sb.append("// IDataSet dataSet = connection.createDataSet((String[]) tables.toArray(new String[0]));\n");
sb.append("// return dataSet;\n");
sb.append("// } catch (DatabaseUnitException | SQLException var6) {\n");
sb.append("// String message = String.format(\"テストケース実施対象テーブル[%s]のテーブルセット作成に失敗しました。\", tables.toString());\n");
sb.append("// throw new RuntimeException(message, var6);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static Map<String, String[]> buildTestCaseExpectEclusiveTableColumnMap(UTContext context)\n");
sb.append("// {\n");
sb.append("// Map<String, String[]> exclusiveColumnMap = new HashMap();\n");
sb.append("// Map<String, TableDefinitionBean> tableDefineMap = context.getTableDefineMap();\n");
sb.append("// Iterator var3 = tableDefineMap.entrySet().iterator();\n");
sb.append("\n");
sb.append("// while (var3.hasNext()) {\n");
sb.append("// Map.Entry<String, TableDefinitionBean> entry = (Map.Entry) var3.next();\n");
sb.append("// TableDefinitionBean bean = (TableDefinitionBean) entry.getValue();\n");
sb.append("// List<String> list = new ArrayList();\n");
sb.append("// List<ColumnDefinitionBean> columns = bean.getTableStructure();\n");
sb.append("// Iterator var8 = columns.iterator();\n");
sb.append("\n");
sb.append("// while (var8.hasNext()) {\n");
sb.append("// ColumnDefinitionBean column = (ColumnDefinitionBean) var8.next();\n");
sb.append("// if (!column.isSubjectOfEntry()) {\n");
sb.append("// list.add(column.getPhysicalName());\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// exclusiveColumnMap.put((String) entry.getKey(), (String[]) list.toArray(new String[0]));\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return exclusiveColumnMap;\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static ITable buildTableData(UTContext context, TableDefinitionBean tableDefinition,\n");
sb.append("// TableDataSetBean tableDataSet) {\n");
sb.append("// try {\n");
sb.append("// String tableName = tableDefinition.getPhysicalTableName();\n");
sb.append("// DefaultTable table = new DefaultTable(tableName, createTableColumnDefine(context,\n");
sb.append("// tableDefinition));\n");
sb.append("// List<Map<String, Object>> dataSetList = tableDataSet.getTableData();\n");
sb.append("// int recNumber = 0;\n");
sb.append("// Iterator var7 = dataSetList.iterator();\n");
sb.append("\n");
sb.append("// while (var7.hasNext()) {\n");
sb.append("// Map<String, Object> dataSetMap = (Map) var7.next();\n");
sb.append("// ++recNumber;\n");
sb.append("// if (logger.isInfoEnabled()) {\n");
sb.append("// String message = String.format(\"DBテーブル登録 テーブル名[%s] レコード番号[%d]\", tableName, recNumber);\n");
sb.append("// logger.info(message);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// Object[] recordData = createTableRecordData(context, tableDefinition, dataSetMap);\n");
sb.append("// table.addRow(recordData);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return table;\n");
sb.append("// } catch (DataSetException var10) {\n");
sb.append("// String message = String.format(\"テーブル[%s]のレコードセット作成に失敗しました。\",\n");
sb.append("// tableDefinition.getPhysicalTableName());\n");
sb.append("// throw new RuntimeException(message, var10);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static void outputPrepareTable(UTContext context, IDataSet dataSet) {\n");
sb.append("// try {\n");
sb.append("// String[] var2 = dataSet.getTableNames();\n");
sb.append("// int var24 = var2.length;\n");
sb.append("\n");
sb.append("// for (int var4 = 0; var4 < var24; ++var4) {\n");
sb.append("// String tableName = var2[var4];\n");
sb.append("\n");
sb.append("// String message;\n");
sb.append("// try {\n");
sb.append("// tableName = tableName.toUpperCase();\n");
sb.append("// ITable tableSet = dataSet.getTable(tableName);\n");
sb.append("// TableDefinitionBean tableDefinitionBean = (TableDefinitionBean)\n");
sb.append("// context.getTableDefineMap().get(tableName);\n");
sb.append("// TableDataSetBean tableDataSetBean = new TableDataSetBean();\n");
sb.append("// tableDataSetBean.setPhysicalTableName(tableName);\n");
sb.append("// List<Map<String, Object>> tableData = new ArrayList();\n");
sb.append("// tableDataSetBean.setTableData(tableData);\n");
sb.append("// List<ColumnDefinitionBean> columnDefList = tableDefinitionBean.getTableStructure();\n");
sb.append("// int rowCount = tableSet.getRowCount();\n");
sb.append("\n");
sb.append("// for (int i = 0; i < rowCount; ++i) {\n");
sb.append("// Map<String, Object> recordData = new LinkedHashMap();\n");
sb.append("// Iterator var14 = columnDefList.iterator();\n");
sb.append("\n");
sb.append("// while (var14.hasNext()) {\n");
sb.append("// ColumnDefinitionBean colDef = (ColumnDefinitionBean) var14.next();\n");
sb.append("// String colName = colDef.getPhysicalName();\n");
sb.append("// Object value = tableSet.getValue(i, colName);\n");
sb.append("// DataType dataType = getTableColumnDataType(colDef.getDataType());\n");
sb.append("// Object convValue = getColumnValueForJson(colDef, dataType, value);\n");
sb.append("// recordData.put(colName, convValue);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// tableData.add(recordData);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// ObjectMapper mapper = new ObjectMapper();\n");
sb.append("// mapper.enable(DeserializationFeature.USE_BIG_DECIMAL_FOR_FLOATS);\n");
sb.append("// File fileObj = new File(getResultDbOutputFileName(context, tableName, \"01_prepare\"));\n");
sb.append("// mapper.writerWithDefaultPrettyPrinter().writeValue(fileObj, tableDataSetBean);\n");
sb.append("// } catch (DataSetException var20) {\n");
sb.append("// message = String.format(\"テーブル[%s]のレコードセット取得に失敗しました。\", tableName);\n");
sb.append("// throw new RuntimeException(message, var20);\n");
sb.append("// } catch (JsonProcessingException var21) {\n");
sb.append("// message = String.format(\"テーブル[%s]のレコードセットからJSON変換に失敗しました。\", tableName);\n");
sb.append("// throw new RuntimeException(message, var21);\n");
sb.append("// } catch (IOException var22) {\n");
sb.append("// message = String.format(\"テーブル[%s]をファイルへの出力に失敗しました。\", tableName);\n");
sb.append("// throw new RuntimeException(message, var22);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// } catch (DataSetException var23) {\n");
sb.append("// String message = String.format(\"準備用レコードセット取得に失敗しました。\");\n");
sb.append("// throw new RuntimeException(message, var23);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static void outputResultTable(UTContext context, ITable tableSet, String tableName, String\n");
sb.append("// fileFooter) {\n");
sb.append("// String message;\n");
sb.append("// try {\n");
sb.append("// TableDefinitionBean tableDefinitionBean = (TableDefinitionBean)\n");
sb.append("// context.getTableDefineMap().get(tableName);\n");
sb.append("// TableDataSetBean tableDataSetBean = new TableDataSetBean();\n");
sb.append("// tableDataSetBean.setPhysicalTableName(tableName);\n");
sb.append("// List<Map<String, Object>> tableData = new ArrayList();\n");
sb.append("// tableDataSetBean.setTableData(tableData);\n");
sb.append("// List<ColumnDefinitionBean> columnDefList = tableDefinitionBean.getTableStructure();\n");
sb.append("// int rowCount = tableSet.getRowCount();\n");
sb.append("\n");
sb.append("// for (int i = 0; i < rowCount; ++i) {\n");
sb.append("// Map<String, Object> recordData = new LinkedHashMap();\n");
sb.append("// Iterator var11 = columnDefList.iterator();\n");
sb.append("\n");
sb.append("// while (var11.hasNext()) {\n");
sb.append("// ColumnDefinitionBean colDef = (ColumnDefinitionBean) var11.next();\n");
sb.append("// String colName = colDef.getPhysicalName();\n");
sb.append("// Object value = tableSet.getValue(i, colName);\n");
sb.append("// DataType dataType = getTableColumnDataType(colDef.getDataType());\n");
sb.append("// Object convValue = getColumnValueForJson(colDef, dataType, value);\n");
sb.append("// recordData.put(colName, convValue);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// tableData.add(recordData);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// ObjectMapper mapper = new ObjectMapper();\n");
sb.append("// mapper.enable(DeserializationFeature.USE_BIG_DECIMAL_FOR_FLOATS);\n");
sb.append("// File fileObj = new File(getResultDbOutputFileName(context, tableName, fileFooter));\n");
sb.append("// mapper.writerWithDefaultPrettyPrinter().writeValue(fileObj, tableDataSetBean);\n");
sb.append("// } catch (DataSetException var17) {\n");
sb.append("// message = String.format(\"テーブル[%s]のレコードセット取得に失敗しました。\", tableName);\n");
sb.append("// throw new RuntimeException(message, var17);\n");
sb.append("// } catch (JsonProcessingException var18) {\n");
sb.append("// message = String.format(\"テーブル[%s]のレコードセットからJSON変換に失敗しました。\", tableName);\n");
sb.append("// throw new RuntimeException(message, var18);\n");
sb.append("// } catch (IOException var19) {\n");
sb.append("// message = String.format(\"テーブル[%s]をファイルへの出力に失敗しました。\", tableName);\n");
sb.append("// throw new RuntimeException(message, var19);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static String getResultDbOutputFileName(UTContext context, String tableName, String\n");
sb.append("// footer) {\n");
sb.append("// String directory = UTFolderUtil.getResultDbTestCaseOutputFolder(context);\n");
sb.append("// (new File(directory)).mkdirs();\n");
sb.append("// String result = String.format(\"%s/%s_%s.json\", directory, tableName, footer);\n");
sb.append("// return result;\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static void assertTableDataEquals(UTContext context, ITable expectedTable, ITable\n");
sb.append("// actualTable, String[] ignoreColumns) {\n");
sb.append("// ITable filteredExpected = null;\n");
sb.append("// ITable filteredActual = null;\n");
sb.append("\n");
sb.append("// try {\n");
sb.append("// ITableMetaData metaData = expectedTable.getTableMetaData();\n");
sb.append("// List<String> includedColumns = (List)\n");
sb.append("// Arrays.stream(metaData.getColumns()).map(Column::getColumnName).map(String::toLowerCase).filter((col)\n");
sb.append("// -> {\n");
sb.append("// boolean var2;\n");
sb.append("// if (ignoreColumns != null) {\n");
sb.append("// Stream var10000 = Arrays.stream(ignoreColumns).map(String::toLowerCase);\n");
sb.append("// Objects.requireNonNull(col);\n");
sb.append("// if (!var10000.noneMatch(col::equals)) {\n");
sb.append("// var2 = false;\n");
sb.append("// return var2;\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// var2 = true;\n");
sb.append("// return var2;\n");
sb.append("// }).collect(Collectors.toList());\n");
sb.append("// String[] comparedColumns = (String[]) includedColumns.toArray(new String[0]);\n");
sb.append("// String[] sortedColumns = (String[]) includedColumns.toArray(new String[0]);\n");
sb.append("// SortedTable sortedExpected = new SortedTable(expectedTable, sortedColumns);\n");
sb.append("// SortedTable sortedActual = new SortedTable(new UpperCaseTable(actualTable), sortedColumns);\n");
sb.append("// filteredExpected = new ColumnSelectingTable(sortedExpected, comparedColumns);\n");
sb.append("// filteredActual = new ColumnSelectingTable(sortedActual, comparedColumns);\n");
sb.append("// String tableName = expectedTable.getTableMetaData().getTableName().toUpperCase();\n");
sb.append("// outputResultTable(context, filteredExpected, tableName, \"02_expect\");\n");
sb.append("// outputResultTable(context, filteredActual, tableName, \"03_actual\");\n");
sb.append("// dumpTable(context, filteredExpected, 0);\n");
sb.append("// dumpTable(context, filteredActual, 1);\n");
sb.append("// } catch (Exception var13) {\n");
sb.append("// String message = String.format(\"テーブル比較の準備処理で想定外の異常が発生しました。テストケースID[%s]\",\n");
sb.append("// context.getTestCaseId());\n");
sb.append("// throw new RuntimeException(message);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// try {\n");
sb.append("// DiffCollectingFailureHandler failureHandler = new DiffCollectingFailureHandler();\n");
sb.append("// DbUnitAssert dbUnitAssert = new DbUnitAssert();\n");
sb.append("// dbUnitAssert.assertEquals(filteredExpected, filteredActual, failureHandler);\n");
sb.append("// if (failureHandler.getDiffList().isEmpty()) {\n");
sb.append("// String message = String.format(\"テストケース[%s]テーブル比較結果=【差分なし】\", context.getTestCaseId());\n");
sb.append("// logger.info(message);\n");
sb.append("// } else {\n");
sb.append("// StringBuilder sb = new StringBuilder();\n");
sb.append("// sb.append(\"テストケース[\").append(context.getTestCaseId()).append(\"]\");\n");
sb.append("// sb.append(\"テーブル比較結果=【★★★差分あり★★★】\").append(System.lineSeparator());\n");
sb.append("// sb.append(\" - 【テーブル期待差分：項目差分一覧】\").append(System.lineSeparator());\n");
sb.append("// Iterator var21 = failureHandler.getDiffList().iterator();\n");
sb.append("\n");
sb.append("// while (var21.hasNext()) {\n");
sb.append("// Object diff = var21.next();\n");
sb.append("// if (diff instanceof Difference) {\n");
sb.append("// Difference dbDiff = (Difference) diff;\n");
sb.append("// sb.append(\" - Row[\").append(dbDiff.getRowIndex()).append(\"], \").append(\"結果項目[\")\n");
sb.append("// .append(dbDiff.getActualTable().getTableMetaData().getTableName()).append(\".\").append(dbDiff.getColumnName()).append(\"]=[\")\n");
sb.append("// .append(dbDiff.getActualValue()).append(\"],\n");
sb.append("// \").append(\"期待項目[\").append(dbDiff.getExpectedTable().getTableMetaData().getTableName())\n");
sb.append("// .append(\".\").append(dbDiff.getColumnName()).append(\"]=[\").append(dbDiff.getExpectedValue()).append(\"],\n");
sb.append("// \");\n");
sb.append("// if (dbDiff.getFailMessage().length() > 0) {\n");
sb.append("// sb.append(\"差異内容[\").append(dbDiff.getFailMessage()).append(\"]\");\n");
sb.append("// } else {\n");
sb.append("// sb.append(\"差異内容[項目値不一致]\");\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// sb.append(System.lineSeparator());\n");
sb.append("// } else {\n");
sb.append("// sb.append(diff.toString()).append(System.lineSeparator());\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// logger.error(sb.toString());\n");
sb.append("// context.setDbAssertionError(true);\n");
sb.append("// }\n");
sb.append("// } catch (Throwable var14) {\n");
sb.append("// StringBuilder sb = new StringBuilder();\n");
sb.append("// sb.append(\"テストケース[\").append(context.getTestCaseId()).append(\"]\");\n");
sb.append("// sb.append(\"【テーブル比較結果】差分あり\").append(System.lineSeparator());\n");
sb.append("// sb.append(\" - 【テーブル期待差分：レコード差異】\").append(System.lineSeparator());\n");
sb.append("// sb.append(\" - 差異内容[\").append(var14.getMessage()).append(\"]\").append(System.lineSeparator());\n");
sb.append("// logger.error(sb.toString());\n");
sb.append("// context.setDbAssertionError(true);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// private static Column[] createTableColumnDefine(UTContext context, TableDefinitionBean\n");
sb.append("// tableDefinition) {\n");
sb.append("// List<Column> columnList = new ArrayList();\n");
sb.append("// List<ColumnDefinitionBean> columnDefinitionList = tableDefinition.getTableStructure();\n");
sb.append("// Iterator var4 = columnDefinitionList.iterator();\n");
sb.append("\n");
sb.append("// while (var4.hasNext()) {\n");
sb.append("// ColumnDefinitionBean columnBean = (ColumnDefinitionBean) var4.next();\n");
sb.append("// DataType dataType = getTableColumnDataType(columnBean.getDataType());\n");
sb.append("// Column.Nullable nullable = Column.NULLABLE;\n");
sb.append("// if (columnBean.isNotNull()) {\n");
sb.append("// nullable = Column.NO_NULLS;\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// Column columnData = new Column(columnBean.getPhysicalName(), dataType, nullable);\n");
sb.append("// columnList.add(columnData);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return (Column[]) columnList.toArray(new Column[0]);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// private static Object[] createTableRecordData(UTContext context, TableDefinitionBean\n");
sb.append("// tableDefinition, Map<String, Object> dataSetMap) {\n");
sb.append("// List<Object> recordDataList = new ArrayList();\n");
sb.append("// List<ColumnDefinitionBean> columnDefinitionList = tableDefinition.getTableStructure();\n");
sb.append("// Iterator var5 = columnDefinitionList.iterator();\n");
sb.append("\n");
sb.append("// while (var5.hasNext()) {\n");
sb.append("// ColumnDefinitionBean columnBean = (ColumnDefinitionBean) var5.next();\n");
sb.append("// DataType dataType = getTableColumnDataType(columnBean.getDataType());\n");
sb.append("// Object value = getColumnValueForDb(columnBean, dataType,\n");
sb.append("// dataSetMap.get(columnBean.getPhysicalName()));\n");
sb.append("// recordDataList.add(value);\n");
sb.append("// if (logger.isInfoEnabled()) {\n");
sb.append("// String message = String.format(\"DBテーブル項目設定 テーブル名[%s], 項目名[%s], 値[%s]\",\n");
sb.append("// tableDefinition.getLogicalTableName(), columnBean.getLogicalName(),\n");
sb.append("// value != null ? value.toString() : \"(null)\");\n");
sb.append("// logger.info(message);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return recordDataList.toArray(new Object[0]);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// private static boolean isEnableValue(DataType dataType, Object value) {\n");
sb.append("// if (value == null) {\n");
sb.append("// return false;\n");
sb.append("// } else if (value instanceof String && \"\".equals(value)) {\n");
sb.append("// return DataType.CHAR == dataType || DataType.VARCHAR == dataType || DataType.NCHAR == dataType;\n");
sb.append("// } else {\n");
sb.append("// return true;\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// private static Object getColumnValueForDb(ColumnDefinitionBean columnBean, DataType dataType,\n");
sb.append("// Object srcValue) {\n");
sb.append("// Object destValue = srcValue;\n");
sb.append("// if (!isEnableValue(dataType, srcValue)) {\n");
sb.append("// if (columnBean.isNotNull()) {\n");
sb.append("// if (!columnBean.isSubjectOfEntry()) {\n");
sb.append("// destValue = createTemporaryValue(columnBean, dataType);\n");
sb.append("// }\n");
sb.append("// } else {\n");
sb.append("// destValue = null;\n");
sb.append("// }\n");
sb.append("// } else if (DataType.DATE == dataType) {\n");
sb.append("// logger.info(\"&&&===\" + JsonUtils.toJson(columnBean.getLogicalName()));\n");
sb.append("// logger.info(\"&&&===\" + JsonUtils.toJson(columnBean.getPhysicalName()));\n");
sb.append("// destValue = parseToLocalDate(srcValue.toString()).toString();\n");
sb.append("// } else if (DataType.TIMESTAMP == dataType) {\n");
sb.append("// destValue = toLocalDateTimeToString(parseToLocalDateTime(srcValue.toString()));\n");
sb.append("// } else if (DataType.NUMERIC != dataType && DataType.INTEGER != dataType && DataType.CHAR ==\n");
sb.append("// dataType) {\n");
sb.append("// long dataLength = columnBean.getNumberOfDigits().longValue();\n");
sb.append("// StringBuilder builder = new StringBuilder();\n");
sb.append("// builder.append(srcValue);\n");
sb.append("\n");
sb.append("// for (int i = builder.length(); (long) i < dataLength; ++i) {\n");
sb.append("// builder.append(' ');\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// destValue = builder.toString();\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return destValue;\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// private static String toLocalDateTimeToString(LocalDateTime localDateTime) {\n");
sb.append("// int year = localDateTime.getYear();\n");
sb.append("// int month = localDateTime.getMonthValue();\n");
sb.append("// int day = localDateTime.getDayOfMonth();\n");
sb.append("// int hour = localDateTime.getHour();\n");
sb.append("// int min = localDateTime.getMinute();\n");
sb.append("// int sec = localDateTime.getSecond();\n");
sb.append("// String str = String.format(\"%04d-%02d-%02d %02d:%02d:%02d\", year, month, day, hour, min, sec);\n");
sb.append("// return str;\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// private static Object getColumnValueForJson(ColumnDefinitionBean columnBean, DataType dataType,\n");
sb.append("// Object srcValue) {\n");
sb.append("// Object destValue = srcValue;\n");
sb.append("// if (DataType.DATE == dataType) {\n");
sb.append("// destValue = convertDateFormat(srcValue);\n");
sb.append("// } else if (DataType.TIMESTAMP == dataType) {\n");
sb.append("// destValue = convertDateTimeFormat(srcValue);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return destValue;\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// private static Object createTemporaryValue(ColumnDefinitionBean columnBean, DataType dataType) {\n");
sb.append("// BigDecimal dataLength = columnBean.getNumberOfDigits();\n");
sb.append("// Object value = null;\n");
sb.append("// if (DataType.VARCHAR == dataType) {\n");
sb.append("// value = getTemporaryAlphabetStringValue(dataLength);\n");
sb.append("// } else if (DataType.CHAR == dataType) {\n");
sb.append("// value = getTemporaryAlphabetStringValue(dataLength);\n");
sb.append("// } else if (DataType.DATE == dataType) {\n");
sb.append("// value = getTodayAsString();\n");
sb.append("// } else if (DataType.TIMESTAMP == dataType) {\n");
sb.append("// value = getCurrentTimestampString();\n");
sb.append("// } else if (DataType.NUMERIC == dataType) {\n");
sb.append("// value = getTemporaryNumericValue(dataLength);\n");
sb.append("// } else if (DataType.INTEGER == dataType) {\n");
sb.append("// value = getTemporaryIntegerValue(dataLength);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return value;\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static String getTemporaryAlphabetStringValue(BigDecimal length) {\n");
sb.append("// int intLen = length.intValue();\n");
sb.append("// StringBuilder builder = new StringBuilder(intLen);\n");
sb.append("\n");
sb.append("// for (int i = 0; i < intLen; ++i) {\n");
sb.append("// char randomAlphabet = (char) (65 + RANDOM.nextInt(26));\n");
sb.append("// builder.append(randomAlphabet);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return builder.toString();\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static BigDecimal getTemporaryNumericValue(BigDecimal digit) {\n");
sb.append("// String strDigit = digit.toPlainString();\n");
sb.append("// int precisionPart = 0;\n");
sb.append("// int scalePart = 0;\n");
sb.append("// int dotPos = strDigit.indexOf(46);\n");
sb.append("// if (dotPos > 0) {\n");
sb.append("// try {\n");
sb.append("// precisionPart = Integer.parseInt(strDigit.substring(0, dotPos));\n");
sb.append("// } catch (Exception var10) {\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// try {\n");
sb.append("// scalePart = Integer.parseInt(strDigit.substring(dotPos + 1));\n");
sb.append("// } catch (Exception var9) {\n");
sb.append("// }\n");
sb.append("// } else {\n");
sb.append("// try {\n");
sb.append("// precisionPart = Integer.parseInt(strDigit);\n");
sb.append("// } catch (Exception var8) {\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// int integerPart = precisionPart > scalePart ? precisionPart - scalePart : precisionPart;\n");
sb.append("// StringBuilder nineBuilder = new StringBuilder();\n");
sb.append("\n");
sb.append("// int i;\n");
sb.append("// for (i = 0; i < integerPart; ++i) {\n");
sb.append("// nineBuilder.append('9');\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// if (scalePart > 0) {\n");
sb.append("// nineBuilder.append('.');\n");
sb.append("\n");
sb.append("// for (i = 0; i < scalePart; ++i) {\n");
sb.append("// nineBuilder.append('9');\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return new BigDecimal(nineBuilder.toString());\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static BigDecimal getTemporaryIntegerValue(BigDecimal digit) {\n");
sb.append("// int intLen = digit.intValue();\n");
sb.append("// StringBuilder nineBuilder = new StringBuilder();\n");
sb.append("// if (intLen > 0) {\n");
sb.append("// for (int i = 0; i < intLen; ++i) {\n");
sb.append("// nineBuilder.append('9');\n");
sb.append("// }\n");
sb.append("// } else {\n");
sb.append("// Random random = new Random();\n");
sb.append("// int value = random.nextInt(Integer.MAX_VALUE);\n");
sb.append("// nineBuilder.append(value);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return new BigDecimal(nineBuilder.toString());\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static LocalDate parseToLocalDate(String dateStr) {\n");
sb.append("// try {\n");
sb.append("// return LocalDate.parse(dateStr, DATE_FORMATTER_FOR_CHECK_YYYYMD);\n");
sb.append("// } catch (DateTimeParseException var2) {\n");
sb.append("// throw new IllegalArgumentException(\"日付形式が不正です。期待される形式: yyyy/MM/dd → 入力値: \" + dateStr, var2);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static String convertDateFormat(Object inputDate) {\n");
sb.append("// if (inputDate != null) {\n");
sb.append("// try {\n");
sb.append("// LocalDate date = LocalDate.parse(inputDate.toString(), DATE_FORMATTER_FOR_BUILD);\n");
sb.append("// return date.format(DATE_FORMATTER_FOR_CHECK_YYYYMMDD);\n");
sb.append("// } catch (Exception var2) {\n");
sb.append("// return null;\n");
sb.append("// }\n");
sb.append("// } else {\n");
sb.append("// return null;\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static String getTodayAsString() {\n");
sb.append("// LocalDate today = LocalDate.now();\n");
sb.append("// return today.format(DATE_FORMATTER_FOR_BUILD);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static LocalDateTime parseToLocalDateTime(String timestampStr) {\n");
sb.append("// try {\n");
sb.append("// return LocalDateTime.parse(timestampStr, TIMESTAMP_FORMATTER_CHECK_YYYYMDHMS);\n");
sb.append("// } catch (DateTimeParseException var2) {\n");
sb.append("// throw new IllegalArgumentException(\"日付時刻の形式が不正です。期待される形式: yyyy/MM/dd HH:mm:ss → 入力: \" +\n");
sb.append("// timestampStr, var2);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static String convertDateTimeFormat(Object inputDateTime) {\n");
sb.append("// if (inputDateTime != null) {\n");
sb.append("// String strTimestamp = inputDateTime.toString();\n");
sb.append("// int index = strTimestamp.indexOf(46);\n");
sb.append("// if (index >= 0) {\n");
sb.append("// strTimestamp = strTimestamp.substring(0, index);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// try {\n");
sb.append("// LocalDateTime dateTime = LocalDateTime.parse(strTimestamp, TIMESTAMP_FORMATTER_BUILD);\n");
sb.append("// return dateTime.format(TIMESTAMP_FORMATTER_CHECK_YYYYMMDDHHMMSS);\n");
sb.append("// } catch (Exception var4) {\n");
sb.append("// return null;\n");
sb.append("// }\n");
sb.append("// } else {\n");
sb.append("// return null;\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static String getCurrentTimestampString() {\n");
sb.append("// LocalDateTime now = LocalDateTime.now();\n");
sb.append("// return now.format(TIMESTAMP_FORMATTER_BUILD);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// private static DataType getTableColumnDataType(String strDataType) {\n");
sb.append("// strDataType = strDataType.toLowerCase();\n");
sb.append("// DataType dataType = DataType.VARCHAR;\n");
sb.append("// if (\"varchar\".equals(strDataType)) {\n");
sb.append("// dataType = DataType.VARCHAR;\n");
sb.append("// } else if (\"char\".equals(strDataType)) {\n");
sb.append("// dataType = DataType.CHAR;\n");
sb.append("// } else if (\"date\".equals(strDataType)) {\n");
sb.append("// dataType = DataType.DATE;\n");
sb.append("// } else if (\"timestamp\".equals(strDataType)) {\n");
sb.append("// dataType = DataType.TIMESTAMP;\n");
sb.append("// } else if (\"numeric\".equals(strDataType)) {\n");
sb.append("// dataType = DataType.NUMERIC;\n");
sb.append("// } else if (\"integer\".equals(strDataType)) {\n");
sb.append("// dataType = DataType.INTEGER;\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// return dataType;\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// private static String getBaseName(String fileName) {\n");
sb.append("// int dotIndex = fileName.lastIndexOf(46);\n");
sb.append("// return dotIndex > 0 && dotIndex < fileName.length() - 1 ? fileName.substring(0, dotIndex) :\n");
sb.append("// fileName;\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static void dumpTable(UTContext context, ITable table, int type) throws DataSetException {\n");
sb.append("// ITableMetaData meta = table.getTableMetaData();\n");
sb.append("// StringBuilder sb = new StringBuilder();\n");
sb.append("// sb.append(\"テストケース[\").append(context.getTestCaseId()).append(\"]\");\n");
sb.append("// sb.append(\"テーブルデータダンプ出力\").append(\" 【データ種別】=[\").append(type == 0 ? \"期待値]\" :\n");
sb.append("// \"実行結果]\").append(System.lineSeparator());\n");
sb.append("// sb.append(\" - 【テーブル名】\").append(meta.getTableName()).append(System.lineSeparator());\n");
sb.append("// sb.append(\" - 【カラム定義】\");\n");
sb.append("// Column[] var5 = meta.getColumns();\n");
sb.append("// int i = var5.length;\n");
sb.append("\n");
sb.append("// for (int var7 = 0; var7 < i; ++var7) {\n");
sb.append("// Column col = var5[var7];\n");
sb.append("// sb.append(col.getColumnName()).append(\"(\").append(col.getDataType()).append(\"), \");\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// sb.append(System.lineSeparator());\n");
sb.append("// int rowCount = table.getRowCount();\n");
sb.append("// sb.append(\" - 【レコード数】\").append(rowCount);\n");
sb.append("\n");
sb.append("// for (i = 0; i < rowCount; ++i) {\n");
sb.append("// sb.append(System.lineSeparator()).append(\" - Row[\").append(i).append(\"]=\");\n");
sb.append("// Column[] var13 = meta.getColumns();\n");
sb.append("// int var14 = var13.length;\n");
sb.append("\n");
sb.append("// for (int var9 = 0; var9 < var14; ++var9) {\n");
sb.append("// Column col = var13[var9];\n");
sb.append("// Object val = table.getValue(i, col.getColumnName());\n");
sb.append("// sb.append(\"[\").append(val).append(\"], \");\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// logger.info(sb.toString());\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// public static void loadIgnoreTestCaseDb(UTContext context) {\n");
sb.append("// Map<String, Map<String, Set<String>>> result = new HashMap();\n");
sb.append("// File file = new File(\"./test/ignore_testcase_db_id.txt\");\n");
sb.append("// if (!file.exists()) {\n");
sb.append("// context.setIgnoreTestCaseTableMap(result);\n");
sb.append("// } else {\n");
sb.append("// try {\n");
sb.append("// BufferedReader br = new BufferedReader(new FileReader(file));\n");
sb.append("\n");
sb.append("// String line;\n");
sb.append("// try {\n");
sb.append("// while ((line = br.readLine()) != null) {\n");
sb.append("// line = line.trim();\n");
sb.append("// if (!line.isEmpty() && !line.startsWith(\"#\")) {\n");
sb.append("// String[] cols = line.split(\",\", -1);\n");
sb.append("// if (cols.length == 3) {\n");
sb.append("// String testCaseId = cols[0];\n");
sb.append("// String tableName = cols[1].toUpperCase();\n");
sb.append("// String category = cols[2].toLowerCase();\n");
sb.append("// Map<String, Set<String>> tableMap = (Map) result.computeIfAbsent(testCaseId, (k) -> {\n");
sb.append("// return new HashMap();\n");
sb.append("// });\n");
sb.append("// Set<String> categories = (Set) tableMap.computeIfAbsent(tableName, (k) -> {\n");
sb.append("// return new HashSet();\n");
sb.append("// });\n");
sb.append("// categories.add(category);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("// } catch (Throwable var12) {\n");
sb.append("// try {\n");
sb.append("// br.close();\n");
sb.append("// } catch (Throwable var11) {\n");
sb.append("// var12.addSuppressed(var11);\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// throw var12;\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// br.close();\n");
sb.append("// } catch (Exception var13) {\n");
sb.append("// }\n");
sb.append("\n");
sb.append("// context.setIgnoreTestCaseTableMap(result);\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("// }\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
return sb.toString();
    }
    /**
     * 
     */
    public String getATest() {
StringBuilder sb=new StringBuilder();
sb.append("package " + Sub.class.getPackage().getName()+  ".sub;\n");
sb.append("\n");
sb.append("\n");
sb.append("import org.junit.platform.engine.discovery.DiscoverySelectors;\n");
sb.append("import org.junit.platform.launcher.Launcher;\n");
sb.append("import org.junit.platform.launcher.LauncherDiscoveryRequest;\n");
sb.append("import org.junit.platform.launcher.core.LauncherDiscoveryRequestBuilder;\n");
sb.append("import org.junit.platform.launcher.core.LauncherFactory;\n");
sb.append("import org.junit.platform.launcher.listeners.SummaryGeneratingListener;\n");
sb.append("import org.junit.platform.launcher.listeners.TestExecutionSummary;\n");
sb.append("\n");
sb.append("import ch.qos.logback.classic.LoggerContext;\n");
sb.append("import ch.qos.logback.classic.joran.JoranConfigurator;\n");
sb.append("import jp.or.jasdec.sbf.cp.ma.batch.biz.CPMAB203Tester;\n");
sb.append("import jp.or.jasdec.sbf.cp.ma.batch.biz.Sub;\n");
sb.append("import org.slf4j.Logger;\n");
sb.append("import org.slf4j.LoggerFactory;\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
sb.append("// 在自动测试入口 src\\test\\java\\jp\\or\\jasdec\\sbf\\cp\\ma\\batch\\biz\\sub\\ATest.java 里，把 String[] arr = {\"000001\"}; 改成合适的测试 case；同时查看 git 最新差分，\n");
sb.append("// 结合 C:\\opt\\bat\\CPMAB203\\test\\inputData\\CPMA\\CPMAB203\\db\\testCase 下的测试数据，补齐某个 case 的 input，让case覆盖此差分\n");
sb.append("//  如果正常系列报异常，要修复，数据差分不用修复\n");
sb.append("\n");
sb.append("\n");
sb.append("//修复 C:\\opt\\bat\\CPMAB203\\test\\inputData\\CPMA\\CPMAB203\\telegram\\testCase 中的 系统异常数据 只修input\n");
sb.append("\n");
sb.append("// 修正000001  中的 测试数据 保证测试\n");
sb.append("// 能夠覆蓋到C:\\opt\\bat\\CPMAB203\\src\\main\\java\\jp\\or\\jasdec\\sbf\\cp\\ma\\batch\\biz\\CPMAB203.java 中的 logger.info(\"===@@@===\"+JsonUtils.toJson(1));\n");
sb.append("\n");
sb.append("// 修正000001  中的 expect数据 让数据没有差分\n");
sb.append("\n");
sb.append("\n");
sb.append("public class ATest {\n");
sb.append("  // implementation 'org.junit.platform:junit-platform-launcher:1.9.1'\n");
sb.append("  public static void main(String[] artgs) {\n");
sb.append("     Sub.as.clearLog();\n");
sb.append("    test();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  public static void test() {\n");
sb.append("\n");
sb.append("    Class<?> testClass = CPMAB203Tester.class;\n");
sb.append("\n");
sb.append("    // 000001 直接用于确认发行业者保证信息检查是否进入 1/3/5 分支。\n");
sb.append("    String[] arr = {\"000001\"};\n");
sb.append("    for (String caseNum : arr) {\n");
sb.append("      auttest(caseNum, testClass);\n");
sb.append("    }\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  public static void auttest(String caseNum, Class<?> testClass) {\n");
sb.append("    // 000001 需要保留预先准备好的 telegram 输入参数，避免初始化时把呼出元类型清空后提前系统异常。\n");
sb.append("    if (!\"000001\".equals(caseNum)) {\n");
sb.append("      Sub.as.createInitTelegramExpectInputNull(caseNum);\n");
sb.append("    }\n");
sb.append("    setOneCase(caseNum, testClass);\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 从一条sql语句设置一个case\n");
sb.append("   */\n");
sb.append("  public static void setOneCase(String caseNum, Class<?> testClass) {\n");
sb.append("    // SubIn.txt放入sql语句 此语句经过了 jsonDataCheck() 保证语句正确后处理\n");
sb.append("    try {\n");
sb.append("      String methodName = \"test\" + caseNum;\n");
sb.append("      // 执行一次放入的语句 查看log日志中是否存在问题\n");
sb.append("      testOneMethod(caseNum, methodName, testClass, 0);\n");
sb.append("    } catch (Exception e) {\n");
sb.append("      e.printStackTrace();\n");
sb.append("    }\n");
sb.append("\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 对一个测试方法进行测试\n");
sb.append("   */\n");
sb.append("  public static void testOneMethod(String caseNum, String methodName, Class<?> testClass, Integer flag) {\n");
sb.append("\n");
sb.append("    SummaryGeneratingListener summaryListener = new SummaryGeneratingListener();\n");
sb.append("\n");
sb.append("    LauncherDiscoveryRequest request =\n");
sb.append("        LauncherDiscoveryRequestBuilder.request().selectors(DiscoverySelectors.selectMethod(testClass, methodName)).build();\n");
sb.append("\n");
sb.append("    Launcher launcher = LauncherFactory.create();\n");
sb.append("    launcher.registerTestExecutionListeners(summaryListener);\n");
sb.append("    launcher.execute(request);\n");
sb.append("\n");
sb.append("    TestExecutionSummary summary = summaryListener.getSummary();\n");
sb.append("    long failures = summary.getFailures().size();\n");
sb.append("\n");
sb.append("    if (failures == 0) {\n");
sb.append("      System.out.println(\"=== \" + methodName + \"测试成功 ===\");\n");
sb.append("    } else {\n");
sb.append("\n");
sb.append("      System.out.println(\"=== 测试失败 \" + failures + \" 个 ===\");\n");
sb.append("      String path = Sub.u.getProjectDir(\"log\") + \"app.log\";\n");
sb.append("      String content = Sub.u.readFile(path);\n");
sb.append("\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 重新配置log防止写文件时总是原来的数据\n");
sb.append("   */\n");
sb.append("  public static void forceResetLogback() {\n");
sb.append("\n");
sb.append("    LoggerContext context = (LoggerContext) LoggerFactory.getILoggerFactory();\n");
sb.append("    try {\n");
sb.append("      // 1. 停止 logback\n");
sb.append("      context.stop();\n");
sb.append("\n");
sb.append("      // 2. 读取 logback.xml 配置重新加载\n");
sb.append("      JoranConfigurator configurator = new JoranConfigurator();\n");
sb.append("      configurator.setContext(context);\n");
sb.append("\n");
sb.append("      // 清理旧配置\n");
sb.append("      context.reset();\n");
sb.append("\n");
sb.append("      // ⭐⭐ 使用你的 logback.xml 路径\n");
sb.append("      configurator.doConfigure(Sub.as.getSrcTestResourcesPath() + \"logback.xml\");\n");
sb.append("\n");
sb.append("      System.out.println(\"Logback 已重新加载配置\");\n");
sb.append("    } catch (Exception e) {\n");
sb.append("      e.printStackTrace();\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("}\n");
sb.append("\n");
sb.append("\n");
sb.append("\n");
return sb.toString();
    }
    /**
     * 
     */
    public String getDbXlsmImportTool() {
StringBuilder sb=new StringBuilder();
sb.append("package " + Sub.class.getPackage().getName()+  ".sub;\n");
sb.append("\n");
sb.append("import java.io.ByteArrayInputStream;\n");
sb.append("import java.io.IOException;\n");
sb.append("import java.io.InputStream;\n");
sb.append("import java.nio.charset.StandardCharsets;\n");
sb.append("import java.nio.file.Files;\n");
sb.append("import java.nio.file.Path;\n");
sb.append("import java.nio.file.Paths;\n");
sb.append("import java.nio.file.StandardCopyOption;\n");
sb.append("import java.security.MessageDigest;\n");
sb.append("import java.security.NoSuchAlgorithmException;\n");
sb.append("import java.util.ArrayList;\n");
sb.append("import java.util.Arrays;\n");
sb.append("import java.util.Comparator;\n");
sb.append("import java.util.Enumeration;\n");
sb.append("import java.util.HashMap;\n");
sb.append("import java.util.LinkedHashMap;\n");
sb.append("import java.util.LinkedHashSet;\n");
sb.append("import java.util.List;\n");
sb.append("import java.util.Locale;\n");
sb.append("import java.util.Map;\n");
sb.append("import java.util.Set;\n");
sb.append("import java.util.TreeMap;\n");
sb.append("import java.util.function.Predicate;\n");
sb.append("import java.util.regex.Matcher;\n");
sb.append("import java.util.regex.Pattern;\n");
sb.append("import java.util.zip.CRC32;\n");
sb.append("import java.util.zip.ZipEntry;\n");
sb.append("import java.util.zip.ZipFile;\n");
sb.append("import java.util.zip.ZipOutputStream;\n");
sb.append("\n");
sb.append("import javax.xml.XMLConstants;\n");
sb.append("import javax.xml.parsers.DocumentBuilderFactory;\n");
sb.append("import javax.xml.parsers.ParserConfigurationException;\n");
sb.append("\n");
sb.append("import org.w3c.dom.Document;\n");
sb.append("import org.w3c.dom.Element;\n");
sb.append("import org.w3c.dom.NodeList;\n");
sb.append("import org.xml.sax.SAXException;\n");
sb.append("\n");
sb.append("import com.fasterxml.jackson.core.type.TypeReference;\n");
sb.append("import com.fasterxml.jackson.databind.ObjectMapper;\n");
sb.append("\n");
sb.append("/**\n");
sb.append(" * DB 数据生成工具（xlsm）与 testCase JSON 的通用同步工具。\n");
sb.append(" *\n");
sb.append(" * <p>本工具不使用 POI 保存 Workbook。它保留原始 ZIP 部件和 worksheet 外壳，只替换已有\n");
sb.append(" * input/expect 数据页的 {@code <sheetData>}。因此宏、ActiveX、关系、Office 扩展命名空间和\n");
sb.append(" * {@code mc:Ignorable} 不会被通用 XML 序列化器重写。</p>\n");
sb.append(" *\n");
sb.append(" * <p>通用性来自 define JSON 和工作表 C4 的物理表名，不包含任何业务表专用分支。</p>\n");
sb.append(" */\n");
sb.append("public final class DbXlsmImportTool {\n");
sb.append("\n");
sb.append("  private static final ObjectMapper OBJECT_MAPPER = new ObjectMapper();\n");
sb.append("  private static final String WORKSHEET_NAMESPACE =\n");
sb.append("      \"http://schemas.openxmlformats.org/spreadsheetml/2006/main\";\n");
sb.append("  private static final String OFFICE_REL_NAMESPACE =\n");
sb.append("      \"http://schemas.openxmlformats.org/officeDocument/2006/relationships\";\n");
sb.append("  private static final int DATA_START_COLUMN = 11;\n");
sb.append("  private static final int OUTPUT_NUMBER_ROW = 9;\n");
sb.append("  private static final int TEST_CASE_ID_ROW = 10;\n");
sb.append("  private static final int EXCLUDED_ROW = 11;\n");
sb.append("  private static final int OPERATION_ROW = 12;\n");
sb.append("  private static final int HEADER_ROW = 13;\n");
sb.append("  private static final int DATA_START_ROW = 14;\n");
sb.append("  private static final Pattern SHEET_DATA_PATTERN = Pattern.compile(\n");
sb.append("      \"<sheetData(?:\\\\s[^>]*)?>.*?</sheetData>\", Pattern.DOTALL);\n");
sb.append("  private static final Pattern ROW_PATTERN = Pattern.compile(\n");
sb.append("      \"<row\\\\b[^>]*\\\\br=\\\"(\\\\d+)\\\"[^>]*>.*?</row>\", Pattern.DOTALL);\n");
sb.append("  private static final Pattern CELL_REFERENCE_PATTERN = Pattern.compile(\n");
sb.append("      \"\\\\br=\\\"([A-Z]+)(\\\\d+)\\\"\");\n");
sb.append("  private static final Pattern STYLE_PATTERN = Pattern.compile(\"\\\\bs=\\\"([^\\\"]+)\\\"\");\n");
sb.append("\n");
sb.append("  private DbXlsmImportTool() {\n");
sb.append("    // Utility class.\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 无参数时自动发现当前工程的目标 xlsm 与 testCase 根目录；也可显式传入两者。\n");
sb.append("   *\n");
sb.append("   * @param args {@code <xlsmPath> <testCaseRootOrDir> [caseId...]}\n");
sb.append("   * @throws Exception 目录、JSON、OOXML 或写入校验失败时抛出\n");
sb.append("   */\n");
sb.append("  public static void main(String[] args) throws Exception {\n");
sb.append("    Invocation invocation = resolveInvocation(args);\n");
sb.append("    ImportResult result = importCases(invocation.xlsmPath(), invocation.caseDirectories());\n");
sb.append("    System.out.println(\"synced workbook: \" + invocation.xlsmPath());\n");
sb.append("    System.out.println(\"backup: \" + result.backupPath());\n");
sb.append("    System.out.println(\"cases=\" + result.caseCount()\n");
sb.append("        + \", sheets=\" + result.updatedSheetCount()\n");
sb.append("        + \", records=\" + result.recordCount());\n");
sb.append("    System.out.println(\"vbaProject.bin SHA-256=\" + result.vbaSha256());\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  /**\n");
sb.append("   * 单独 Case 的快照导入入口。导入后的数据页只包含该 Case。\n");
sb.append("   *\n");
sb.append("   * @param xlsmPath 目标工作簿\n");
sb.append("   * @param testCaseDir 单独 Case 目录\n");
sb.append("   * @throws IOException 导入失败时抛出\n");
sb.append("   */\n");
sb.append("  public static void importTestCase(Path xlsmPath, Path testCaseDir) throws IOException {\n");
sb.append("    importCases(xlsmPath, List.of(testCaseDir));\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static Invocation resolveInvocation(String[] args) throws IOException {\n");
sb.append("    if (args.length == 1) {\n");
sb.append("      throw new IllegalArgumentException(\n");
sb.append("          \"Usage: DbXlsmImportTool <xlsmPath> <testCaseRootOrDir> [caseId...]\");\n");
sb.append("    }\n");
sb.append("    if (args.length >= 2) {\n");
sb.append("      Path xlsmPath = Paths.get(args[0]).toAbsolutePath().normalize();\n");
sb.append("      Path sourcePath = Paths.get(args[1]).toAbsolutePath().normalize();\n");
sb.append("      return new Invocation(xlsmPath, resolveCaseDirectories(sourcePath, args));\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    Path projectRoot = Paths.get(\"\").toAbsolutePath().normalize();\n");
sb.append("    Path testDirectory = projectRoot.resolve(\"test\");\n");
sb.append("    Path xlsmPath = findSinglePath(testDirectory, 1,\n");
sb.append("        path -> path.getFileName().toString().startsWith(\"020_DB\")\n");
sb.append("            && path.getFileName().toString().endsWith(\"v2.15_template_CP.xlsm\")\n");
sb.append("            && !path.getFileName().toString().contains(\".before\"),\n");
sb.append("        \"target xlsm\");\n");
sb.append("    Path testCaseRoot = findSinglePath(testDirectory.resolve(\"inputData\"), 8,\n");
sb.append("        path -> Files.isDirectory(path)\n");
sb.append("            && \"testCase\".equals(path.getFileName().toString())\n");
sb.append("            && path.getParent() != null\n");
sb.append("            && \"db\".equals(path.getParent().getFileName().toString()),\n");
sb.append("        \"db/testCase directory\");\n");
sb.append("    return new Invocation(xlsmPath, resolveCaseDirectories(testCaseRoot, new String[0]));\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static Path findSinglePath(\n");
sb.append("      Path root, int maxDepth, Predicate<Path> predicate, String description) throws IOException {\n");
sb.append("    if (!Files.isDirectory(root)) {\n");
sb.append("      throw new IOException(\"Directory not found while discovering \" + description + \": \" + root);\n");
sb.append("    }\n");
sb.append("    try (var stream = Files.find(root, maxDepth, (path, attributes) -> predicate.test(path))) {\n");
sb.append("      List<Path> matches = stream.sorted().toList();\n");
sb.append("      if (matches.size() != 1) {\n");
sb.append("        throw new IOException(\"Expected exactly one \" + description + \" under \" + root\n");
sb.append("            + \", but found \" + matches);\n");
sb.append("      }\n");
sb.append("      return matches.get(0).toAbsolutePath().normalize();\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static List<Path> resolveCaseDirectories(Path sourcePath, String[] args) throws IOException {\n");
sb.append("    if (args.length >= 3) {\n");
sb.append("      List<Path> selected = new ArrayList<>();\n");
sb.append("      for (int index = 2; index < args.length; index++) {\n");
sb.append("        Path caseDirectory = sourcePath.resolve(normalizeCaseId(args[index]));\n");
sb.append("        requireCaseDirectory(caseDirectory);\n");
sb.append("        selected.add(caseDirectory);\n");
sb.append("      }\n");
sb.append("      return selected.stream().sorted(Comparator.comparing(DbXlsmImportTool::caseId)).toList();\n");
sb.append("    }\n");
sb.append("    if (isCaseDirectory(sourcePath)) {\n");
sb.append("      return List.of(sourcePath);\n");
sb.append("    }\n");
sb.append("    if (!Files.isDirectory(sourcePath)) {\n");
sb.append("      throw new IOException(\"testCase root not found: \" + sourcePath);\n");
sb.append("    }\n");
sb.append("    try (var stream = Files.list(sourcePath)) {\n");
sb.append("      List<Path> cases = stream.filter(Files::isDirectory)\n");
sb.append("          .filter(DbXlsmImportTool::isCaseDirectory)\n");
sb.append("          .sorted(Comparator.comparing(DbXlsmImportTool::caseId))\n");
sb.append("          .toList();\n");
sb.append("      if (cases.isEmpty()) {\n");
sb.append("        throw new IOException(\"No testCase directory found under: \" + sourcePath);\n");
sb.append("      }\n");
sb.append("      return cases;\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static boolean isCaseDirectory(Path path) {\n");
sb.append("    return Files.isDirectory(path.resolve(\"input\")) && Files.isDirectory(path.resolve(\"expect\"));\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static void requireCaseDirectory(Path path) throws IOException {\n");
sb.append("    if (!isCaseDirectory(path)) {\n");
sb.append("      throw new IOException(\"Invalid testCase directory: \" + path);\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String caseId(Path caseDirectory) {\n");
sb.append("    return normalizeCaseId(caseDirectory.getFileName().toString());\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String normalizeCaseId(String rawCaseId) {\n");
sb.append("    String normalized = rawCaseId == null ? \"\" : rawCaseId.trim();\n");
sb.append("    if (normalized.startsWith(\"testCase\")) {\n");
sb.append("      normalized = normalized.substring(\"testCase\".length());\n");
sb.append("    }\n");
sb.append("    if (!normalized.matches(\"\\\\d{6}\")) {\n");
sb.append("      throw new IllegalArgumentException(\"testCaseId must be 6 digits: \" + rawCaseId);\n");
sb.append("    }\n");
sb.append("    return normalized;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static ImportResult importCases(Path xlsmPath, List<Path> caseDirectories) throws IOException {\n");
sb.append("    if (!Files.isRegularFile(xlsmPath)) {\n");
sb.append("      throw new IOException(\"xlsm not found: \" + xlsmPath);\n");
sb.append("    }\n");
sb.append("    if (caseDirectories.isEmpty()) {\n");
sb.append("      throw new IOException(\"No test cases selected.\");\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    Path caseRoot = caseDirectories.get(0).getParent();\n");
sb.append("    for (Path caseDirectory : caseDirectories) {\n");
sb.append("      requireCaseDirectory(caseDirectory);\n");
sb.append("      if (!caseRoot.equals(caseDirectory.getParent())) {\n");
sb.append("        throw new IOException(\"All test cases must have the same root: \" + caseDirectories);\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("    Path defineDirectory = caseRoot.getParent().resolve(\"define\");\n");
sb.append("    Map<String, List<String>> definitions = readDefinitions(defineDirectory);\n");
sb.append("    LinkedHashMap<String, ZipPart> parts = readZip(xlsmPath);\n");
sb.append("    byte[] macroBefore = requiredPart(parts, \"xl/vbaProject.bin\").bytes();\n");
sb.append("    List<String> sharedStrings = readSharedStrings(parts.get(\"xl/sharedStrings.xml\"));\n");
sb.append("    Set<String> registeredCaseIds = readRegisteredCaseIds(parts, sharedStrings);\n");
sb.append("    List<Path> activeCaseDirectories = caseDirectories.stream()\n");
sb.append("        .filter(path -> registeredCaseIds.contains(caseId(path)))\n");
sb.append("        .toList();\n");
sb.append("    if (activeCaseDirectories.isEmpty()) {\n");
sb.append("      throw new IOException(\"No selected test case is registered in テストケース一覧: \" + caseDirectories);\n");
sb.append("    }\n");
sb.append("    Map<SheetKey, String> dataSheets = discoverDataSheets(parts, sharedStrings);\n");
sb.append("\n");
sb.append("    int updatedSheetCount = 0;\n");
sb.append("    int recordCount = 0;\n");
sb.append("    Set<String> updatedParts = new LinkedHashSet<>();\n");
sb.append("    for (Map.Entry<SheetKey, String> entry : dataSheets.entrySet()) {\n");
sb.append("      SheetKey key = entry.getKey();\n");
sb.append("      List<String> definedFields = definitions.get(key.physicalTableName());\n");
sb.append("      if (definedFields == null) {\n");
sb.append("        throw new IOException(\"define JSON not found for sheet \" + key + \".\");\n");
sb.append("      }\n");
sb.append("      String partName = entry.getValue();\n");
sb.append("      ZipPart originalPart = requiredPart(parts, partName);\n");
sb.append("      List<String> worksheetFields = readWorksheetFields(originalPart.bytes(), sharedStrings);\n");
sb.append("      if (!worksheetFields.equals(definedFields)) {\n");
sb.append("        throw new IOException(\"Worksheet definition differs from define JSON: \" + key\n");
sb.append("            + \"\\nworksheet=\" + worksheetFields + \"\\ndefine=\" + definedFields);\n");
sb.append("      }\n");
sb.append("\n");
sb.append("      List<CaseRecord> records = readCaseRecords(activeCaseDirectories, key, definedFields);\n");
sb.append("      byte[] rebuilt = rebuildSheetData(originalPart.bytes(), records, definedFields);\n");
sb.append("      parts.put(partName, originalPart.withBytes(rebuilt));\n");
sb.append("      updatedParts.add(partName);\n");
sb.append("      updatedSheetCount++;\n");
sb.append("      recordCount += records.size();\n");
sb.append("    }\n");
sb.append("    validateAllJsonTablesHaveSheets(activeCaseDirectories, dataSheets.keySet());\n");
sb.append("\n");
sb.append("    Path backupPath = backupPath(xlsmPath);\n");
sb.append("    Files.copy(xlsmPath, backupPath, StandardCopyOption.REPLACE_EXISTING);\n");
sb.append("    Path temporary = Files.createTempFile(xlsmPath.getParent(), \"db-xlsm-import-\", \".xlsm\");\n");
sb.append("    try {\n");
sb.append("      writeZip(temporary, parts);\n");
sb.append("      validateWrittenWorkbook(temporary, parts.keySet(), updatedParts, macroBefore);\n");
sb.append("      Files.move(temporary, xlsmPath, StandardCopyOption.REPLACE_EXISTING);\n");
sb.append("    } finally {\n");
sb.append("      Files.deleteIfExists(temporary);\n");
sb.append("    }\n");
sb.append("    return new ImportResult(\n");
sb.append("        backupPath, activeCaseDirectories.size(), updatedSheetCount, recordCount, sha256Hex(macroBefore));\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static Set<String> readRegisteredCaseIds(\n");
sb.append("      Map<String, ZipPart> parts, List<String> sharedStrings) throws IOException {\n");
sb.append("    Map<String, String> sheetParts = discoverSheetParts(parts);\n");
sb.append("    String listPartName = sheetParts.get(\"テストケース一覧\");\n");
sb.append("    if (listPartName == null) {\n");
sb.append("      throw new IOException(\"テストケース一覧 sheet not found.\");\n");
sb.append("    }\n");
sb.append("    Document document = parseXml(requiredPart(parts, listPartName).bytes());\n");
sb.append("    Set<String> result = new LinkedHashSet<>();\n");
sb.append("    NodeList cells = document.getElementsByTagNameNS(WORKSHEET_NAMESPACE, \"c\");\n");
sb.append("    for (int index = 0; index < cells.getLength(); index++) {\n");
sb.append("      Element cell = (Element) cells.item(index);\n");
sb.append("      Matcher matcher = Pattern.compile(\"B(\\\\d+)\").matcher(cell.getAttribute(\"r\"));\n");
sb.append("      if (!matcher.matches() || Integer.parseInt(matcher.group(1)) < 7) {\n");
sb.append("        continue;\n");
sb.append("      }\n");
sb.append("      String value = elementCellValue(cell, sharedStrings).trim();\n");
sb.append("      if (value.startsWith(\"testCase\")) {\n");
sb.append("        value = value.substring(\"testCase\".length());\n");
sb.append("      }\n");
sb.append("      if (value.matches(\"\\\\d{6}\")) {\n");
sb.append("        result.add(value);\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("    if (result.isEmpty()) {\n");
sb.append("      throw new IOException(\"No testCaseId found in テストケース一覧.\");\n");
sb.append("    }\n");
sb.append("    return result;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static Map<String, List<String>> readDefinitions(Path defineDirectory) throws IOException {\n");
sb.append("    if (!Files.isDirectory(defineDirectory)) {\n");
sb.append("      throw new IOException(\"define directory not found: \" + defineDirectory);\n");
sb.append("    }\n");
sb.append("    Map<String, List<String>> definitions = new HashMap<>();\n");
sb.append("    try (var stream = Files.list(defineDirectory)) {\n");
sb.append("      for (Path jsonFile : stream.filter(path -> path.toString().endsWith(\".json\")).sorted().toList()) {\n");
sb.append("        Map<String, Object> json = readJsonObject(jsonFile);\n");
sb.append("        String table = String.valueOf(json.getOrDefault(\"physicalTableName\", \"\"));\n");
sb.append("        List<String> fields = new ArrayList<>();\n");
sb.append("        Object structure = json.get(\"tableStructure\");\n");
sb.append("        if (structure instanceof List<?> fieldList) {\n");
sb.append("          for (Object rawField : fieldList) {\n");
sb.append("            if (rawField instanceof Map<?, ?> field) {\n");
sb.append("              fields.add(String.valueOf(field.get(\"physicalName\")));\n");
sb.append("            }\n");
sb.append("          }\n");
sb.append("        }\n");
sb.append("        if (!table.isBlank() && !fields.isEmpty()) {\n");
sb.append("          definitions.put(table, List.copyOf(fields));\n");
sb.append("        }\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("    return definitions;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static Map<String, Object> readJsonObject(Path jsonFile) throws IOException {\n");
sb.append("    return OBJECT_MAPPER.readValue(Files.readString(jsonFile, StandardCharsets.UTF_8),\n");
sb.append("        new TypeReference<LinkedHashMap<String, Object>>() {});\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static List<CaseRecord> readCaseRecords(\n");
sb.append("      List<Path> caseDirectories, SheetKey key, List<String> fields) throws IOException {\n");
sb.append("    List<CaseRecord> records = new ArrayList<>();\n");
sb.append("    String directoryName = \"inp\".equals(key.mode()) ? \"input\" : \"expect\";\n");
sb.append("    for (Path caseDirectory : caseDirectories) {\n");
sb.append("      Path jsonFile = caseDirectory.resolve(directoryName).resolve(key.physicalTableName() + \".json\");\n");
sb.append("      if (!Files.isRegularFile(jsonFile)) {\n");
sb.append("        continue;\n");
sb.append("      }\n");
sb.append("      Map<String, Object> payload = readJsonObject(jsonFile);\n");
sb.append("      String payloadTable = String.valueOf(payload.getOrDefault(\"physicalTableName\", \"\"));\n");
sb.append("      if (!key.physicalTableName().equals(payloadTable)) {\n");
sb.append("        throw new IOException(\"physicalTableName mismatch: \" + jsonFile);\n");
sb.append("      }\n");
sb.append("      Object rawTableData = payload.get(\"tableData\");\n");
sb.append("      if (!(rawTableData instanceof List<?> tableData)) {\n");
sb.append("        throw new IOException(\"tableData must be an array: \" + jsonFile);\n");
sb.append("      }\n");
sb.append("      for (int rowIndex = 0; rowIndex < tableData.size(); rowIndex++) {\n");
sb.append("        Object rawRecord = tableData.get(rowIndex);\n");
sb.append("        if (!(rawRecord instanceof Map<?, ?> sourceRecord)) {\n");
sb.append("          throw new IOException(\"tableData record must be an object: \" + jsonFile + \"[\" + rowIndex + \"]\");\n");
sb.append("        }\n");
sb.append("        LinkedHashMap<String, Object> record = new LinkedHashMap<>();\n");
sb.append("        for (Map.Entry<?, ?> value : sourceRecord.entrySet()) {\n");
sb.append("          record.put(String.valueOf(value.getKey()), value.getValue());\n");
sb.append("        }\n");
sb.append("        if (!new ArrayList<>(record.keySet()).equals(fields)) {\n");
sb.append("          throw new IOException(\"JSON field order/set differs from define JSON: \" + jsonFile\n");
sb.append("              + \"[\" + rowIndex + \"]\\njson=\" + record.keySet() + \"\\ndefine=\" + fields);\n");
sb.append("        }\n");
sb.append("        records.add(new CaseRecord(caseId(caseDirectory), record));\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("    return records;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static void validateAllJsonTablesHaveSheets(\n");
sb.append("      List<Path> caseDirectories, Set<SheetKey> sheetKeys) throws IOException {\n");
sb.append("    Set<SheetKey> available = Set.copyOf(sheetKeys);\n");
sb.append("    for (Path caseDirectory : caseDirectories) {\n");
sb.append("      for (String directoryName : List.of(\"input\", \"expect\")) {\n");
sb.append("        String mode = \"input\".equals(directoryName) ? \"inp\" : \"exp\";\n");
sb.append("        try (var stream = Files.list(caseDirectory.resolve(directoryName))) {\n");
sb.append("          for (Path jsonFile : stream.filter(path -> path.toString().endsWith(\".json\")).toList()) {\n");
sb.append("            String table = String.valueOf(readJsonObject(jsonFile).getOrDefault(\"physicalTableName\", \"\"));\n");
sb.append("            SheetKey key = new SheetKey(mode, table);\n");
sb.append("            if (!available.contains(key)) {\n");
sb.append("              throw new IOException(\"Workbook data sheet not found for \" + key + \": \" + jsonFile);\n");
sb.append("            }\n");
sb.append("          }\n");
sb.append("        }\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static Map<SheetKey, String> discoverDataSheets(\n");
sb.append("      Map<String, ZipPart> parts, List<String> sharedStrings) throws IOException {\n");
sb.append("    Map<String, String> sheetParts = discoverSheetParts(parts);\n");
sb.append("    Map<SheetKey, String> result = new LinkedHashMap<>();\n");
sb.append("    for (Map.Entry<String, String> sheetEntry : sheetParts.entrySet()) {\n");
sb.append("      String sheetName = sheetEntry.getKey();\n");
sb.append("      String mode = sheetName.startsWith(\"inp_\") ? \"inp\" : sheetName.startsWith(\"exp_\") ? \"exp\" : \"\";\n");
sb.append("      if (mode.isBlank()) {\n");
sb.append("        continue;\n");
sb.append("      }\n");
sb.append("      String partName = sheetEntry.getValue();\n");
sb.append("      ZipPart worksheetPart = requiredPart(parts, partName);\n");
sb.append("      String physicalTableName = readCellValue(worksheetPart.bytes(), \"C4\", sharedStrings);\n");
sb.append("      if (!physicalTableName.isBlank()) {\n");
sb.append("        SheetKey key = new SheetKey(mode, physicalTableName);\n");
sb.append("        if (result.put(key, partName) != null) {\n");
sb.append("          throw new IOException(\"Duplicate data sheet for \" + key);\n");
sb.append("        }\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("    return result;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static Map<String, String> discoverSheetParts(Map<String, ZipPart> parts) throws IOException {\n");
sb.append("    Document workbook = parseXml(requiredPart(parts, \"xl/workbook.xml\").bytes());\n");
sb.append("    Document relationships = parseXml(requiredPart(parts, \"xl/_rels/workbook.xml.rels\").bytes());\n");
sb.append("    Map<String, String> targets = new HashMap<>();\n");
sb.append("    NodeList relationshipNodes = relationships.getElementsByTagNameNS(\"*\", \"Relationship\");\n");
sb.append("    for (int index = 0; index < relationshipNodes.getLength(); index++) {\n");
sb.append("      Element relation = (Element) relationshipNodes.item(index);\n");
sb.append("      targets.put(relation.getAttribute(\"Id\"), normalizePartName(\"xl\", relation.getAttribute(\"Target\")));\n");
sb.append("    }\n");
sb.append("\n");
sb.append("    Map<String, String> result = new LinkedHashMap<>();\n");
sb.append("    NodeList sheets = workbook.getElementsByTagNameNS(WORKSHEET_NAMESPACE, \"sheet\");\n");
sb.append("    for (int index = 0; index < sheets.getLength(); index++) {\n");
sb.append("      Element sheet = (Element) sheets.item(index);\n");
sb.append("      String sheetName = sheet.getAttribute(\"name\");\n");
sb.append("      String relationId = sheet.getAttributeNS(OFFICE_REL_NAMESPACE, \"id\");\n");
sb.append("      String partName = targets.get(relationId);\n");
sb.append("      if (partName == null || !parts.containsKey(partName)) {\n");
sb.append("        throw new IOException(\"Worksheet relationship target missing: \" + sheetName + \" -> \" + partName);\n");
sb.append("      }\n");
sb.append("      result.put(sheetName, partName);\n");
sb.append("    }\n");
sb.append("    return result;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String normalizePartName(String base, String target) {\n");
sb.append("    if (target.startsWith(\"/\")) {\n");
sb.append("      return target.substring(1);\n");
sb.append("    }\n");
sb.append("    Path normalized = Paths.get(base).resolve(target.replace('/', java.io.File.separatorChar)).normalize();\n");
sb.append("    return normalized.toString().replace(java.io.File.separatorChar, '/');\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static List<String> readSharedStrings(ZipPart part) throws IOException {\n");
sb.append("    if (part == null) {\n");
sb.append("      return List.of();\n");
sb.append("    }\n");
sb.append("    Document document = parseXml(part.bytes());\n");
sb.append("    NodeList items = document.getElementsByTagNameNS(WORKSHEET_NAMESPACE, \"si\");\n");
sb.append("    List<String> result = new ArrayList<>();\n");
sb.append("    for (int index = 0; index < items.getLength(); index++) {\n");
sb.append("      result.add(descendantText((Element) items.item(index), \"t\"));\n");
sb.append("    }\n");
sb.append("    return result;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String readCellValue(byte[] worksheet, String reference, List<String> sharedStrings)\n");
sb.append("      throws IOException {\n");
sb.append("    Document document = parseXml(worksheet);\n");
sb.append("    NodeList cells = document.getElementsByTagNameNS(WORKSHEET_NAMESPACE, \"c\");\n");
sb.append("    for (int index = 0; index < cells.getLength(); index++) {\n");
sb.append("      Element cell = (Element) cells.item(index);\n");
sb.append("      if (reference.equals(cell.getAttribute(\"r\"))) {\n");
sb.append("        return elementCellValue(cell, sharedStrings);\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("    return \"\";\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static List<String> readWorksheetFields(byte[] worksheet, List<String> sharedStrings)\n");
sb.append("      throws IOException {\n");
sb.append("    Document document = parseXml(worksheet);\n");
sb.append("    TreeMap<Integer, String> fields = new TreeMap<>();\n");
sb.append("    NodeList cells = document.getElementsByTagNameNS(WORKSHEET_NAMESPACE, \"c\");\n");
sb.append("    for (int index = 0; index < cells.getLength(); index++) {\n");
sb.append("      Element cell = (Element) cells.item(index);\n");
sb.append("      Matcher matcher = Pattern.compile(\"D(\\\\d+)\").matcher(cell.getAttribute(\"r\"));\n");
sb.append("      if (matcher.matches()) {\n");
sb.append("        int row = Integer.parseInt(matcher.group(1));\n");
sb.append("        String value = elementCellValue(cell, sharedStrings);\n");
sb.append("        if (row >= DATA_START_ROW && !value.isBlank()) {\n");
sb.append("          fields.put(row, value);\n");
sb.append("        }\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("    List<String> result = new ArrayList<>();\n");
sb.append("    int expectedRow = DATA_START_ROW;\n");
sb.append("    for (Map.Entry<Integer, String> field : fields.entrySet()) {\n");
sb.append("      if (field.getKey() != expectedRow++) {\n");
sb.append("        throw new IOException(\"Non-contiguous worksheet field rows: \" + fields.keySet());\n");
sb.append("      }\n");
sb.append("      result.add(field.getValue());\n");
sb.append("    }\n");
sb.append("    return result;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String elementCellValue(Element cell, List<String> sharedStrings) throws IOException {\n");
sb.append("    String type = cell.getAttribute(\"t\");\n");
sb.append("    if (\"inlineStr\".equals(type)) {\n");
sb.append("      return descendantText(cell, \"t\");\n");
sb.append("    }\n");
sb.append("    String value = descendantText(cell, \"v\");\n");
sb.append("    if (\"s\".equals(type) && !value.isBlank()) {\n");
sb.append("      int index = Integer.parseInt(value);\n");
sb.append("      if (index < 0 || index >= sharedStrings.size()) {\n");
sb.append("        throw new IOException(\"Invalid shared string index: \" + index);\n");
sb.append("      }\n");
sb.append("      return sharedStrings.get(index);\n");
sb.append("    }\n");
sb.append("    return value;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String descendantText(Element parent, String localName) {\n");
sb.append("    NodeList nodes = parent.getElementsByTagNameNS(WORKSHEET_NAMESPACE, localName);\n");
sb.append("    StringBuilder result = new StringBuilder();\n");
sb.append("    for (int index = 0; index < nodes.getLength(); index++) {\n");
sb.append("      result.append(nodes.item(index).getTextContent());\n");
sb.append("    }\n");
sb.append("    return result.toString();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static byte[] rebuildSheetData(\n");
sb.append("      byte[] worksheetBytes, List<CaseRecord> records, List<String> fields) throws IOException {\n");
sb.append("    String worksheet = new String(worksheetBytes, StandardCharsets.UTF_8);\n");
sb.append("    Matcher sheetDataMatcher = SHEET_DATA_PATTERN.matcher(worksheet);\n");
sb.append("    if (!sheetDataMatcher.find()) {\n");
sb.append("      throw new IOException(\"sheetData not found in worksheet.\");\n");
sb.append("    }\n");
sb.append("    String originalSheetData = sheetDataMatcher.group();\n");
sb.append("    Matcher rowMatcher = ROW_PATTERN.matcher(originalSheetData);\n");
sb.append("    StringBuffer rebuiltSheetData = new StringBuffer();\n");
sb.append("    Set<Integer> requiredRows = new LinkedHashSet<>(\n");
sb.append("        List.of(OUTPUT_NUMBER_ROW, TEST_CASE_ID_ROW, EXCLUDED_ROW, OPERATION_ROW, HEADER_ROW));\n");
sb.append("    for (int index = 0; index < fields.size(); index++) {\n");
sb.append("      requiredRows.add(DATA_START_ROW + index);\n");
sb.append("    }\n");
sb.append("    Set<Integer> foundRows = new LinkedHashSet<>();\n");
sb.append("    while (rowMatcher.find()) {\n");
sb.append("      int rowNumber = Integer.parseInt(rowMatcher.group(1));\n");
sb.append("      String originalRow = rowMatcher.group();\n");
sb.append("      String style = templateStyle(originalRow);\n");
sb.append("      String rowWithoutData = removeDataCells(originalRow);\n");
sb.append("      if (requiredRows.contains(rowNumber)) {\n");
sb.append("        foundRows.add(rowNumber);\n");
sb.append("        StringBuilder cells = new StringBuilder();\n");
sb.append("        for (int recordIndex = 0; recordIndex < records.size(); recordIndex++) {\n");
sb.append("          String reference = columnName(DATA_START_COLUMN + recordIndex) + rowNumber;\n");
sb.append("          Object value;\n");
sb.append("          if (rowNumber == OUTPUT_NUMBER_ROW) {\n");
sb.append("            value = recordIndex + 1;\n");
sb.append("          } else if (rowNumber == TEST_CASE_ID_ROW) {\n");
sb.append("            value = records.get(recordIndex).caseId();\n");
sb.append("          } else if (rowNumber == EXCLUDED_ROW || rowNumber == OPERATION_ROW || rowNumber == HEADER_ROW) {\n");
sb.append("            value = null;\n");
sb.append("          } else {\n");
sb.append("            String field = fields.get(rowNumber - DATA_START_ROW);\n");
sb.append("            value = records.get(recordIndex).record().get(field);\n");
sb.append("          }\n");
sb.append("          cells.append(cellXml(reference, style, value));\n");
sb.append("        }\n");
sb.append("        rowWithoutData = rowWithoutData.replace(\"</row>\", cells + \"</row>\");\n");
sb.append("      }\n");
sb.append("      rowMatcher.appendReplacement(rebuiltSheetData, Matcher.quoteReplacement(rowWithoutData));\n");
sb.append("    }\n");
sb.append("    rowMatcher.appendTail(rebuiltSheetData);\n");
sb.append("    if (!foundRows.equals(requiredRows)) {\n");
sb.append("      Set<Integer> missing = new LinkedHashSet<>(requiredRows);\n");
sb.append("      missing.removeAll(foundRows);\n");
sb.append("      throw new IOException(\"Required worksheet rows are missing: \" + missing);\n");
sb.append("    }\n");
sb.append("    String result = worksheet.substring(0, sheetDataMatcher.start())\n");
sb.append("        + rebuiltSheetData + worksheet.substring(sheetDataMatcher.end());\n");
sb.append("    validateWorksheetNamespaces(result);\n");
sb.append("    byte[] rebuiltBytes = result.getBytes(StandardCharsets.UTF_8);\n");
sb.append("    validateCellStructure(rebuiltBytes);\n");
sb.append("    return rebuiltBytes;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String removeDataCells(String rowXml) throws IOException {\n");
sb.append("    StringBuilder result = new StringBuilder(rowXml.length());\n");
sb.append("    int cursor = 0;\n");
sb.append("    for (CellFragment cell : cellsInRow(rowXml)) {\n");
sb.append("      result.append(rowXml, cursor, cell.start());\n");
sb.append("      if (cell.column() < DATA_START_COLUMN) {\n");
sb.append("        result.append(cell.xml());\n");
sb.append("      }\n");
sb.append("      cursor = cell.end();\n");
sb.append("    }\n");
sb.append("    result.append(rowXml, cursor, rowXml.length());\n");
sb.append("    return result.toString();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String templateStyle(String rowXml) throws IOException {\n");
sb.append("    String fallback = \"\";\n");
sb.append("    for (CellFragment cell : cellsInRow(rowXml)) {\n");
sb.append("      Matcher styleMatcher = STYLE_PATTERN.matcher(cell.xml());\n");
sb.append("      String style = styleMatcher.find() ? styleMatcher.group(1) : \"\";\n");
sb.append("      if (cell.column() == DATA_START_COLUMN) {\n");
sb.append("        return style;\n");
sb.append("      }\n");
sb.append("      if (cell.column() >= DATA_START_COLUMN && fallback.isBlank()) {\n");
sb.append("        fallback = style;\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("    return fallback;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static List<CellFragment> cellsInRow(String rowXml) throws IOException {\n");
sb.append("    List<CellFragment> cells = new ArrayList<>();\n");
sb.append("    int searchFrom = 0;\n");
sb.append("    while (true) {\n");
sb.append("      int start = rowXml.indexOf(\"<c \", searchFrom);\n");
sb.append("      if (start < 0) {\n");
sb.append("        return cells;\n");
sb.append("      }\n");
sb.append("      int openingEnd = rowXml.indexOf('>', start);\n");
sb.append("      if (openingEnd < 0) {\n");
sb.append("        throw new IOException(\"Cell opening tag is incomplete.\");\n");
sb.append("      }\n");
sb.append("      boolean selfClosing = rowXml.charAt(openingEnd - 1) == '/';\n");
sb.append("      int end;\n");
sb.append("      if (selfClosing) {\n");
sb.append("        end = openingEnd + 1;\n");
sb.append("      } else {\n");
sb.append("        int closingStart = rowXml.indexOf(\"</c>\", openingEnd + 1);\n");
sb.append("        if (closingStart < 0) {\n");
sb.append("          throw new IOException(\"Cell closing tag is missing.\");\n");
sb.append("        }\n");
sb.append("        end = closingStart + \"</c>\".length();\n");
sb.append("      }\n");
sb.append("      String cellXml = rowXml.substring(start, end);\n");
sb.append("      Matcher referenceMatcher = CELL_REFERENCE_PATTERN.matcher(cellXml);\n");
sb.append("      if (!referenceMatcher.find()) {\n");
sb.append("        throw new IOException(\"Cell reference is missing: \" + cellXml);\n");
sb.append("      }\n");
sb.append("      cells.add(new CellFragment(\n");
sb.append("          start, end, columnNumber(referenceMatcher.group(1)), cellXml));\n");
sb.append("      searchFrom = end;\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String cellXml(String reference, String style, Object value) {\n");
sb.append("    String styleAttribute = style.isBlank() ? \"\" : \" s=\\\"\" + xmlAttribute(style) + \"\\\"\";\n");
sb.append("    if (value == null) {\n");
sb.append("      return \"<c r=\\\"\" + reference + \"\\\"\" + styleAttribute + \"/>\";\n");
sb.append("    }\n");
sb.append("    if (value instanceof Boolean bool) {\n");
sb.append("      return \"<c r=\\\"\" + reference + \"\\\"\" + styleAttribute + \" t=\\\"b\\\"><v>\"\n");
sb.append("          + (bool ? \"1\" : \"0\") + \"</v></c>\";\n");
sb.append("    }\n");
sb.append("    if (value instanceof Number number) {\n");
sb.append("      return \"<c r=\\\"\" + reference + \"\\\"\" + styleAttribute + \"><v>\"\n");
sb.append("          + xmlText(number.toString()) + \"</v></c>\";\n");
sb.append("    }\n");
sb.append("    String text = String.valueOf(value);\n");
sb.append("    String preserve = !text.isEmpty()\n");
sb.append("        && (Character.isWhitespace(text.charAt(0)) || Character.isWhitespace(text.charAt(text.length() - 1)))\n");
sb.append("        ? \" xml:space=\\\"preserve\\\"\" : \"\";\n");
sb.append("    String textNode = text.isEmpty() ? \"<t/>\" : \"<t\" + preserve + \">\" + xmlText(text) + \"</t>\";\n");
sb.append("    return \"<c r=\\\"\" + reference + \"\\\"\" + styleAttribute\n");
sb.append("        + \" t=\\\"inlineStr\\\"><is>\" + textNode + \"</is></c>\";\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static int columnNumber(String letters) {\n");
sb.append("    int result = 0;\n");
sb.append("    for (int index = 0; index < letters.length(); index++) {\n");
sb.append("      result = result * 26 + letters.charAt(index) - 'A' + 1;\n");
sb.append("    }\n");
sb.append("    return result;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String columnName(int number) {\n");
sb.append("    StringBuilder result = new StringBuilder();\n");
sb.append("    while (number > 0) {\n");
sb.append("      number--;\n");
sb.append("      result.insert(0, (char) ('A' + number % 26));\n");
sb.append("      number /= 26;\n");
sb.append("    }\n");
sb.append("    return result.toString();\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String xmlText(String value) {\n");
sb.append("    return value.replace(\"&\", \"&amp;\").replace(\"<\", \"&lt;\").replace(\">\", \"&gt;\");\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String xmlAttribute(String value) {\n");
sb.append("    return xmlText(value).replace(\"\\\"\", \"&quot;\").replace(\"'\", \"&apos;\");\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static void validateWorksheetNamespaces(String worksheet) throws IOException {\n");
sb.append("    Matcher rootMatcher = Pattern.compile(\"<worksheet\\\\b[^>]*>\").matcher(worksheet);\n");
sb.append("    if (!rootMatcher.find()) {\n");
sb.append("      throw new IOException(\"worksheet root element not found.\");\n");
sb.append("    }\n");
sb.append("    String root = rootMatcher.group();\n");
sb.append("    Matcher ignorableMatcher = Pattern.compile(\"\\\\bIgnorable=\\\"([^\\\"]+)\\\"\").matcher(root);\n");
sb.append("    if (ignorableMatcher.find()) {\n");
sb.append("      for (String prefix : ignorableMatcher.group(1).split(\"\\\\s+\")) {\n");
sb.append("        if (!root.contains(\"xmlns:\" + prefix + \"=\")) {\n");
sb.append("          throw new IOException(\"mc:Ignorable prefix is not declared: \" + prefix);\n");
sb.append("        }\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("    if (Pattern.compile(\"xmlns:ns\\\\d+=\").matcher(root).find()) {\n");
sb.append("      throw new IOException(\"Auto-generated namespace prefix detected in worksheet root.\");\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static LinkedHashMap<String, ZipPart> readZip(Path xlsmPath) throws IOException {\n");
sb.append("    LinkedHashMap<String, ZipPart> result = new LinkedHashMap<>();\n");
sb.append("    try (ZipFile zip = new ZipFile(xlsmPath.toFile(), StandardCharsets.UTF_8)) {\n");
sb.append("      Enumeration<? extends ZipEntry> entries = zip.entries();\n");
sb.append("      while (entries.hasMoreElements()) {\n");
sb.append("        ZipEntry entry = entries.nextElement();\n");
sb.append("        try (InputStream input = zip.getInputStream(entry)) {\n");
sb.append("          byte[] bytes = input.readAllBytes();\n");
sb.append("          result.put(entry.getName(), new ZipPart(\n");
sb.append("              entry.getName(), entry.getMethod(), entry.getTime(), entry.getComment(), entry.getExtra(), bytes));\n");
sb.append("        }\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("    return result;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static ZipPart requiredPart(Map<String, ZipPart> parts, String name) throws IOException {\n");
sb.append("    ZipPart part = parts.get(name);\n");
sb.append("    if (part == null) {\n");
sb.append("      throw new IOException(\"Required XLSM part not found: \" + name);\n");
sb.append("    }\n");
sb.append("    return part;\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static void writeZip(Path target, LinkedHashMap<String, ZipPart> parts) throws IOException {\n");
sb.append("    try (ZipOutputStream output = new ZipOutputStream(Files.newOutputStream(target), StandardCharsets.UTF_8)) {\n");
sb.append("      for (ZipPart part : parts.values()) {\n");
sb.append("        ZipEntry entry = new ZipEntry(part.name());\n");
sb.append("        entry.setMethod(part.method());\n");
sb.append("        if (part.time() >= 0) {\n");
sb.append("          entry.setTime(part.time());\n");
sb.append("        }\n");
sb.append("        if (part.comment() != null) {\n");
sb.append("          entry.setComment(part.comment());\n");
sb.append("        }\n");
sb.append("        if (part.extra() != null) {\n");
sb.append("          entry.setExtra(part.extra());\n");
sb.append("        }\n");
sb.append("        if (part.method() == ZipEntry.STORED) {\n");
sb.append("          CRC32 crc = new CRC32();\n");
sb.append("          crc.update(part.bytes());\n");
sb.append("          entry.setSize(part.bytes().length);\n");
sb.append("          entry.setCompressedSize(part.bytes().length);\n");
sb.append("          entry.setCrc(crc.getValue());\n");
sb.append("        }\n");
sb.append("        output.putNextEntry(entry);\n");
sb.append("        output.write(part.bytes());\n");
sb.append("        output.closeEntry();\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static void validateWrittenWorkbook(\n");
sb.append("      Path workbook, Set<String> expectedParts, Set<String> updatedParts, byte[] macroBefore) throws IOException {\n");
sb.append("    Set<String> actualParts = new LinkedHashSet<>();\n");
sb.append("    byte[] macroAfter = null;\n");
sb.append("    try (ZipFile zip = new ZipFile(workbook.toFile(), StandardCharsets.UTF_8)) {\n");
sb.append("      Enumeration<? extends ZipEntry> entries = zip.entries();\n");
sb.append("      while (entries.hasMoreElements()) {\n");
sb.append("        ZipEntry entry = entries.nextElement();\n");
sb.append("        actualParts.add(entry.getName());\n");
sb.append("        try (InputStream input = zip.getInputStream(entry)) {\n");
sb.append("          byte[] bytes = input.readAllBytes();\n");
sb.append("          if (updatedParts.contains(entry.getName())) {\n");
sb.append("            parseXml(bytes);\n");
sb.append("            validateWorksheetNamespaces(new String(bytes, StandardCharsets.UTF_8));\n");
sb.append("            validateCellStructure(bytes);\n");
sb.append("          }\n");
sb.append("          if (\"xl/vbaProject.bin\".equals(entry.getName())) {\n");
sb.append("            macroAfter = bytes;\n");
sb.append("          }\n");
sb.append("        }\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("    if (!actualParts.equals(expectedParts)) {\n");
sb.append("      throw new IOException(\"ZIP part list changed unexpectedly.\");\n");
sb.append("    }\n");
sb.append("    if (!Arrays.equals(macroBefore, macroAfter)) {\n");
sb.append("      throw new IOException(\"vbaProject.bin changed unexpectedly.\");\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static void validateCellStructure(byte[] worksheetBytes) throws IOException {\n");
sb.append("    Document document = parseXml(worksheetBytes);\n");
sb.append("    NodeList rows = document.getElementsByTagNameNS(WORKSHEET_NAMESPACE, \"row\");\n");
sb.append("    Set<String> workbookReferences = new LinkedHashSet<>();\n");
sb.append("    for (int rowIndex = 0; rowIndex < rows.getLength(); rowIndex++) {\n");
sb.append("      Element row = (Element) rows.item(rowIndex);\n");
sb.append("      int rowNumber = Integer.parseInt(row.getAttribute(\"r\"));\n");
sb.append("      int previousColumn = 0;\n");
sb.append("      NodeList cells = row.getElementsByTagNameNS(WORKSHEET_NAMESPACE, \"c\");\n");
sb.append("      for (int cellIndex = 0; cellIndex < cells.getLength(); cellIndex++) {\n");
sb.append("        Element cell = (Element) cells.item(cellIndex);\n");
sb.append("        String reference = cell.getAttribute(\"r\");\n");
sb.append("        Matcher referenceMatcher = CELL_REFERENCE_PATTERN.matcher(\"r=\\\"\" + reference + \"\\\"\");\n");
sb.append("        if (!referenceMatcher.find()) {\n");
sb.append("          throw new IOException(\"Invalid cell reference: \" + reference);\n");
sb.append("        }\n");
sb.append("        int column = columnNumber(referenceMatcher.group(1));\n");
sb.append("        int referencedRow = Integer.parseInt(referenceMatcher.group(2));\n");
sb.append("        if (referencedRow != rowNumber) {\n");
sb.append("          throw new IOException(\"Cell row does not match parent row: \" + reference);\n");
sb.append("        }\n");
sb.append("        if (!workbookReferences.add(reference)) {\n");
sb.append("          throw new IOException(\"Duplicate cell reference: \" + reference);\n");
sb.append("        }\n");
sb.append("        if (column <= previousColumn) {\n");
sb.append("          throw new IOException(\"Cell references are not strictly ordered at: \" + reference);\n");
sb.append("        }\n");
sb.append("        previousColumn = column;\n");
sb.append("      }\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static Document parseXml(byte[] bytes) throws IOException {\n");
sb.append("    try {\n");
sb.append("      DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();\n");
sb.append("      factory.setNamespaceAware(true);\n");
sb.append("      factory.setFeature(\"http://apache.org/xml/features/disallow-doctype-decl\", true);\n");
sb.append("      factory.setFeature(\"http://xml.org/sax/features/external-general-entities\", false);\n");
sb.append("      factory.setFeature(\"http://xml.org/sax/features/external-parameter-entities\", false);\n");
sb.append("      factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_DTD, \"\");\n");
sb.append("      factory.setAttribute(XMLConstants.ACCESS_EXTERNAL_SCHEMA, \"\");\n");
sb.append("      return factory.newDocumentBuilder().parse(new ByteArrayInputStream(bytes));\n");
sb.append("    } catch (ParserConfigurationException | SAXException exception) {\n");
sb.append("      throw new IOException(\"Invalid OOXML part.\", exception);\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static Path backupPath(Path xlsmPath) {\n");
sb.append("    String fileName = xlsmPath.getFileName().toString();\n");
sb.append("    String stem = fileName.toLowerCase(Locale.ROOT).endsWith(\".xlsm\")\n");
sb.append("        ? fileName.substring(0, fileName.length() - 5) : fileName;\n");
sb.append("    return xlsmPath.resolveSibling(stem + \".before-db-import.xlsm\");\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private static String sha256Hex(byte[] bytes) throws IOException {\n");
sb.append("    try {\n");
sb.append("      byte[] digest = MessageDigest.getInstance(\"SHA-256\").digest(bytes);\n");
sb.append("      StringBuilder result = new StringBuilder();\n");
sb.append("      for (byte value : digest) {\n");
sb.append("        result.append(String.format(\"%02x\", value));\n");
sb.append("      }\n");
sb.append("      return result.toString();\n");
sb.append("    } catch (NoSuchAlgorithmException exception) {\n");
sb.append("      throw new IOException(\"SHA-256 is unavailable.\", exception);\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private record Invocation(Path xlsmPath, List<Path> caseDirectories) {\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private record ImportResult(\n");
sb.append("      Path backupPath, int caseCount, int updatedSheetCount, int recordCount, String vbaSha256) {\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private record SheetKey(String mode, String physicalTableName) {\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private record CaseRecord(String caseId, LinkedHashMap<String, Object> record) {\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private record CellFragment(int start, int end, int column, String xml) {\n");
sb.append("  }\n");
sb.append("\n");
sb.append("  private record ZipPart(\n");
sb.append("      String name, int method, long time, String comment, byte[] extra, byte[] bytes) {\n");
sb.append("\n");
sb.append("    private ZipPart withBytes(byte[] replacement) {\n");
sb.append("      return new ZipPart(name, method, time, comment, extra, replacement);\n");
sb.append("    }\n");
sb.append("  }\n");
sb.append("}\n");
return sb.toString();
    }
}
}

