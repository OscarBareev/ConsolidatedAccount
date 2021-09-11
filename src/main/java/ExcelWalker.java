import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Stream;

public class ExcelWalker {

    private static final int MAX_SIZE = 209715200;


    private int wbCounter = 0;
    String resultWbPath;
    private XSSFWorkbook resultWb;
    private FileOutputStream outStream;
    private FileInputStream inputStream;
    XSSFWorkbook existWb;


    public void doWork(Path path) throws IOException {

        resultWb = new XSSFWorkbook();
        resultWbPath = crtOutPass(path);
        outStream = new FileOutputStream(resultWbPath);
        resultWb.write(outStream);
        outStream.close();


        Stream<Path> filePathStream = Files.walk(path);

        filePathStream
                .filter(Files::isRegularFile)
                .filter(filePath -> filePath.toString().endsWith("xlsx"))
                .forEach(filePath -> {

                    System.out.println(filePath);

                    try {


                        long fileSize = Files.size(Paths.get(resultWbPath));

                        if (fileSize > MAX_SIZE) {

                            resultWbPath = crtOutPass(path);
                            resultWb = new XSSFWorkbook();
                            outStream = new FileOutputStream(resultWbPath);
                            resultWb.write(outStream);
                            outStream.close();

                        }


                        inputStream = new FileInputStream(filePath.toString());
                        existWb = new XSSFWorkbook(inputStream);
                        int sheetsNum = existWb.getNumberOfSheets();

                        for (int i = 0; i < sheetsNum; i++) {

                            System.out.println(i);

                            removeMergedRegions(existWb.getSheetAt(i));
                            CopySheets.copySheets(resultWb.createSheet(), existWb.getSheetAt(i), false);
                        }

                        inputStream.close();

                        outStream = new FileOutputStream(resultWbPath);
                        resultWb.write(outStream);
                        outStream.close();


                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                });
    }

    private String crtOutPass(Path path) {
        wbCounter++;
        return path.getParent().toString() + "\\Таблица счетов " + wbCounter + ".xlsx";
    }

    private void removeMergedRegions(XSSFSheet sheet) {
        if (sheet.getNumMergedRegions() > 0) {
            while (sheet.getNumMergedRegions() > 0) {
                for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                    sheet.removeMergedRegion(i);
                }
            }
        }
    }

}
