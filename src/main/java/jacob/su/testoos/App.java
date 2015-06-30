package jacob.su.testoos;

import java.io.*;
import java.net.ConnectException;
import java.util.Date;
import java.util.concurrent.atomic.AtomicLong;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

/**
 * <p>TODO</p>
 *
 * @author <a href="mailto:ysu2@cisco.com">Yu Su</a>
 * @version 1.0
 */
public class App {


    static final String OpenOffice_HOME = "/opt/openoffice4";
    static final AtomicLong threadCount = new AtomicLong(0);
    static final DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
    static OfficeManager officeManager;
    

    public static void main(String... args) throws InterruptedException {

        configuration.setOfficeHome(new File(OpenOffice_HOME));// 设置OpenOffice.org安装目录
        configuration.setPortNumbers(8100); // 设置转换端口，默认为8100
        configuration.setTaskExecutionTimeout(1000 * 60 * 5L);// 设置任务执行超时为5分钟
        configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);// 设置任务队列超时为24小时
        officeManager = configuration.buildOfficeManager();

        final String currentPath = System.getProperty("user.dir");
        String originalDoc = currentPath + File.separator + "SwitchAndRouterTheory.doc";
        final String originalDocx = currentPath + File.separator + "TestConvertor.docx";
        String originalXLS = currentPath + File.separator + "Workbook2.xls";
        final String originalXLSX = currentPath + File.separator + "Workbook2.xlsx";
        Long totalStart = System.currentTimeMillis();

        for (int i=0; i<100; i++) {
            officeManager.start();
            convertPDF(new File(originalDoc), currentPath);
            convertPDF(new File(originalDocx), currentPath);
            convertPDF(new File(originalXLS), currentPath);
            convertPDF(new File(originalXLSX), currentPath);
            officeManager.stop();
        }
        Long totalEnd = System.currentTimeMillis();
        System.out.println("total time count: "+(totalEnd-totalStart));
        Long threadTotalStart = System.currentTimeMillis();
        officeManager.start();
        for (int i=0; i<100; i++) {
            threadCount.incrementAndGet();
            Thread thread = new Thread(new Runnable() {
                @Override
                public void run() {
                    if ((threadCount.longValue()%2)==0){
                        convertPDF(new File(originalXLSX), currentPath);
                    } else {
                        convertPDF(new File(originalDocx), currentPath);
                    }
                    threadCount.decrementAndGet();
                }
            });
        }
        while (threadCount.longValue() > 0) {
            System.out.println("current thread: "+threadCount);
        }
        officeManager.stop();
        Long threadTotalEnd = System.currentTimeMillis();
        System.out.println("current thread count: "+(threadTotalEnd-threadTotalStart));
        //System.out.println(toHtmlString(new File(originalDoc), currentPath));
    }


    /**
     * 将word文档转换成odf文档
     *
     * @param docFile  需要转换的word文档
     * @param filepath 转换之后html的存放路径
     */
    public static void convertODF(File docFile, String filepath) {

        // 创建保存html的文件
        File pdfFile = new File(filepath + "/" + new Date().getTime()
            + ".odf");


        OfficeManager officeManager = configuration.buildOfficeManager();
        officeManager.start();

        OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
        converter.convert(docFile, pdfFile);

        officeManager.stop();

    }

    /**
     * 将word文档转换成pdf文档
     *
     * @param docFile  需要转换的word文档
     * @param filepath 转换之后html的存放路径
     */
    public static void convertPDF(File docFile, String filepath) {
        Long start = System.currentTimeMillis();
        // 创建保存html的文件
        File pdfFile = new File(filepath + "/" + new Date().getTime()
            + ".pdf");


        OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
        converter.convert(docFile, pdfFile);

        officeManager.stop();

        Long end = System.currentTimeMillis();
        System.out.println("conversion pdf cost: "+(end - start));
    }

    /**
     * 将word文档转换成html文档
     *
     * @param docFile  需要转换的word文档
     * @param filepath 转换之后html的存放路径
     *
     * @return 转换之后的html文件
     */
    public static File convertHtml(File docFile, String filepath) {
        // 创建保存html的文件
        File htmlFile = new File(filepath + "/" + new Date().getTime()
            + ".html");
        DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
        configuration.setOfficeHome(new File(OpenOffice_HOME));// 设置OpenOffice.org安装目录
        configuration.setPortNumbers(8100); // 设置转换端口，默认为8100
        configuration.setTaskExecutionTimeout(1000 * 60 * 5L);// 设置任务执行超时为5分钟
        configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);// 设置任务队列超时为24小时
        OfficeManager officeManager = configuration.buildOfficeManager();
        officeManager.start();

        OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
        converter.convert(docFile, htmlFile);

        officeManager.stop();
        return htmlFile;
    }

    /**
     * 将word转换成html文件，并且获取html文件代码。
     *
     * @param docFile  需要转换的文档
     * @param filepath 文档中图片的保存位置
     *
     * @return 转换成功的html代码
     */
    public static String toHtmlString(File docFile, String filepath) {
        // 转换word文档
        File htmlFile = convertHtml(docFile, filepath);
        // 获取html文件流
        StringBuffer htmlSb = new StringBuffer();
        try {
            BufferedReader br = new BufferedReader(new InputStreamReader(
                new FileInputStream(htmlFile)));
            while (br.ready()) {
                htmlSb.append(br.readLine());
            }
            br.close();
            // 删除临时文件
            htmlFile.delete();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        // HTML文件字符串
        String htmlStr = htmlSb.toString();
        // 返回经过清洁的html文本
        return clearFormat(htmlStr, filepath);
    }

    /**
     * 清除一些不需要的html标记
     *
     * @param htmlStr 带有复杂html标记的html语句
     *
     * @return 去除了不需要html标记的语句
     */
    protected static String clearFormat(String htmlStr, String docImgPath) {
        // 获取body内容的正则
        String bodyReg = "<BODY .*</BODY>";
        Pattern bodyPattern = Pattern.compile(bodyReg);
        Matcher bodyMatcher = bodyPattern.matcher(htmlStr);
        if (bodyMatcher.find()) {
            // 获取BODY内容，并转化BODY标签为DIV
            htmlStr = bodyMatcher.group().replaceFirst("<BODY", "<DIV")
                .replaceAll("</BODY>", "</DIV>");
        }
        // 调整图片地址
        htmlStr = htmlStr.replaceAll("<IMG SRC=\"", "<IMG SRC=\"" + docImgPath
            + "/");
        // 把<P></P>转换成</div></div>保留样式
        // content = content.replaceAll("(<P)([^>]*>.*?)(<\\/P>)",
        // "<div$2</div>");
        // 把<P></P>转换成</div></div>并删除样式
        htmlStr = htmlStr.replaceAll("(<P)([^>]*)(>.*?)(<\\/P>)", "<p$3</p>");
        // 删除不需要的标签
        htmlStr = htmlStr
            .replaceAll(
                "<[/]?(font|FONT|span|SPAN|xml|XML|del|DEL|ins|INS|meta|META|[ovwxpOVWXP]:\\w+)[^>]*?>",
                "");
        // 删除不需要的属性
        htmlStr = htmlStr
            .replaceAll(
                "<([^>]*)(?:lang|LANG|class|CLASS|style|STYLE|size|SIZE|face|FACE|[ovwxpOVWXP]:\\w+)=(?:'[^']*'|\"\"[^\"\"]*\"\"|[^>]+)([^>]*)>",
                "<$1$2>");
        return htmlStr;
    }

}

