/**
 *
 */
package uehara.daishin.exceltsv;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * @author d-uehara
 *
 */
public class ExcelTsv {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		if (2 != args.length){
			System.err.println("excel-tsv ReadDirectory OutputDirectory");
			System.exit(1);
		}
        System.out.println("[INFO]処理を開始しました");
        String arg0=args[0].replaceFirst("\\\\$", "");
        String arg1=args[1].replaceFirst("\\\\$", "");

		Path start_dir = Paths.get(arg0);
		Path abs_start_dir=start_dir.toAbsolutePath();
		String abs_start_str=abs_start_dir.toString();
		int abs_start_len=abs_start_str.length();

		Path output_dir =Paths.get(arg1);
		Path abs_output_dir=output_dir.toAbsolutePath();
		String abs_outoput_str=abs_output_dir.toString();

        try {
			Files.walkFileTree(start_dir, new SimpleFileVisitor<Path>() {
			    @Override
			    public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
			    	String excelFileName=file.getFileName().toString();
			    	if (excelFileName.endsWith(".xlsx")){
				    	Path fullpath=file.toAbsolutePath();
			        	String f1=fullpath.toString();
			        	String excel_relative_file_path=f1.substring(abs_start_len);
			        	String tsv_relative_file_path=excel_relative_file_path.replaceFirst(".xlsx$", "");
				        String to_file_base = abs_outoput_str+tsv_relative_file_path;

				        Path pp=file.getParent(); // 親ディレクトリ
				        Path abs_pp=pp.toAbsolutePath();
				        String abs_pp_path=abs_pp.toString();

			        	String relative_pp_path=abs_pp_path.substring(abs_start_len);

				    	String mkdirpath=abs_outoput_str+relative_pp_path;

				    	Path mkdir_path=Paths.get(mkdirpath);

				    	if(!Files.exists(mkdir_path)){
					    	Files.createDirectories(mkdir_path);
				    	}

				    	try {
							excelToTsv(f1,to_file_base);
						} catch (Exception e) {
							e.printStackTrace();
							System.exit(1);
						}

			    	}

			        return FileVisitResult.CONTINUE;
			    }
			});
		} catch (IOException e) {
			e.printStackTrace();
			System.exit(1);
		}
        System.out.println("[INFO]処理を終了しました");
	}


	private static void excelToTsv(String excel_path, String to_file_base){
        Workbook workbook;
        File excel_file = new File(excel_path);
		// Excelのオープン
        workbook=null;
        try {
			workbook = WorkbookFactory.create(excel_file, null, true);
			excelBookToTsv(workbook,to_file_base);
		} catch (EncryptedDocumentException | IOException e) {
			System.err.println("[ERROR]"+excel_file+"を開けませんでした");
			e.printStackTrace();
			return;
		} finally {
			if (workbook!=null){
				try {
					workbook.close();
				} catch (IOException e) {
					System.err.println("[ERROR]"+excel_file+"を閉じれませんでした");
					e.printStackTrace();
				}
			}
		}

	}

	private static void excelBookToTsv(Workbook workbook, String to_file_base){
		Iterator<Sheet> sheet_iterator=workbook.sheetIterator();
		while(sheet_iterator.hasNext()){
			Sheet sheet=sheet_iterator.next();
			int index = workbook.getSheetIndex(sheet);
			openTsvFileWriter(sheet,index ,to_file_base);
		}
	}

	private static void openTsvFileWriter(Sheet sheet, int index, String to_file_base){

		String tsv_file_path=to_file_base+"."+String.format("%03d", index)+"."+sheet.getSheetName()+".tsv";

		FileWriter tsv_writer=null;
		try {

			tsv_writer = new FileWriter(tsv_file_path);
			openTsvPrintWriter(sheet, tsv_writer);

		} catch (IOException e1) {
			System.err.println("[ERROR]"+tsv_file_path+"を開くのに失敗しました");
			e1.printStackTrace();
		}finally{
			if ( null != tsv_writer ){
				try {
					tsv_writer.close();
				} catch (IOException e2) {
					System.err.println("[ERROR]"+tsv_file_path+"を閉じるのに失敗しました");
				}
			}

		}
		return;

	}

	private static void openTsvPrintWriter(Sheet sheet, FileWriter tsv_fw){
		PrintWriter tsv_print_writer=null;
		tsv_print_writer=new PrintWriter(new BufferedWriter(tsv_fw));

		for (int row_num=0; row_num <= sheet.getLastRowNum(); row_num++){
			Row row=sheet.getRow(row_num);
			if (null != row){
				for ( int cell_num=0; cell_num < row.getLastCellNum(); cell_num++){
					Cell cell = row.getCell(cell_num);
					if ( null != cell){
						tsv_print_writer.print(toTsvString(getCellString(cell)));
					}
					if (cell_num + 1 < row.getLastCellNum()){
						tsv_print_writer.print("\t");
					}
				}
			}
			if (row_num + 1 <= sheet.getLastRowNum()){
				tsv_print_writer.print("\n");
			}
		}

		tsv_print_writer.close();

	}

	private static String getCellString(Cell cell){
		String ret="";
		if ( null == cell ){
			ret= "";
		} else {
			switch(cell.getCellType()){
			case NUMERIC:
				double d=cell.getNumericCellValue();
				if ( d == (long) d){
					ret=String.format("%d",(long)d);
				} else {
					ret=String.format("%s",d);
				}
				break;
			case STRING:
				ret= cell.getStringCellValue();
				break;
			case BLANK:
				ret= "";
				break;
			case BOOLEAN:
				ret=String.valueOf(cell.getBooleanCellValue());
				break;
			case FORMULA:
				ret=getFormulaValue(cell);
				if (ret == null){
					ret="";
				}
				break;
			case ERROR:
				ret="##ERROR="+String.valueOf(cell.getErrorCellValue())+"##";
				break;
			default:
				ret= "";
				break;
			}

		}
		return ret;
	}

	private static String toTsvString(String src){
		return src.replace("\\","\\\\" ).replace("\n","\\n" ).replace("\t","\\t" ).replace("\"","\\\"" );
	}


    private static String getFormulaValue(Cell fcell) {
		String ret="";
        Workbook book = fcell.getSheet().getWorkbook();
        CreationHelper helper = book.getCreationHelper();
        FormulaEvaluator evaluator = helper.createFormulaEvaluator();
        CellValue value = evaluator.evaluate(fcell);

        switch(value.getCellType()){
		case NUMERIC:
			double d=value.getNumberValue();
			if ( d == (long) d){
				ret=String.format("%d",(long)d);
			} else {
				ret=String.format("%s",d);
			}
			break;
		case STRING:
			ret=value.getStringValue();
			break;
		case BLANK:
			ret="";
			break;
		case BOOLEAN:
			ret=String.valueOf(value.getBooleanValue());
			break;
		case FORMULA:
			ret="##ERROR=FORMULA_ERROR##";
			break;
		case ERROR:
			ret="##ERROR="+String.valueOf(value.getErrorValue())+"##";
			break;
		default:
			ret= "";
			break;
		}
        return ret;
    }
}


