import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class TextToExcel {

    public static void main(String[] args) {
        String filename = "BTSolverResultsTable2.txt";
        String workbookName = "BTSolverResults.xls";
        String sheetname = "BTSolverResults";
        List<String> lines = new ArrayList<>();
        int colLength = 10;

        // read file, store in list
        try {
            BufferedReader reader = new BufferedReader(new FileReader(filename));
            String line;
            while ( (line = reader.readLine()) != null) {
                lines.add(line);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        // separate lines into words, removing leading and trailing whitespace
        List<String> results = new ArrayList<>();
        for (String line : lines) {
            String[] words = line.trim().split("\\|");
            for (String word : words) {
                if (word.contains("-") || word.contains("*"))
                    continue;
                results.add(word.trim());
            }
        }

        // create workbook and sheet
        WritableSheet sheet;
        WritableWorkbook workbook;
        int rowLength = results.size() / colLength;  // there are 10 columns per row
        try {
            workbook = Workbook.createWorkbook(new File(workbookName));
            sheet = workbook.createSheet(sheetname, 0);

            int index = 0;
            for (int row = 0; row < rowLength; row++) {
                for (int col = 0; col < colLength; col++) {
                    String content = col == 4 ? results.get(index) + " ms" : col > 6 ? results.get(index) + "%" : results.get(index);
                    sheet.addCell(new Label(col, row, content));
                    index++;
                }
            }
            workbook.write();
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
