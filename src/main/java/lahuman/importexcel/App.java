package lahuman.importexcel;

import java.io.File;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.jdbc.core.JdbcTemplate;

import jxl.Sheet;
import jxl.Workbook;

/**
 * Import Excel Data to Table
 *
 */
@SpringBootApplication
public class App implements CommandLineRunner
{
    public static void main( String[] args )
    {
        SpringApplication.run(App.class, args);
    }

    @Autowired
    JdbcTemplate jdbcTemplate;
    
	public void run(String... arg0) throws Exception {
		Workbook wb = null;
		try {
			wb = Workbook.getWorkbook(new File("EXCEL FILE PATH"));
			Sheet sheet = wb.getSheet(3);
			for(int i=1; i< 200; i++) {
					this.jdbcTemplate.update("insert into table (column1, column2, column3) values (?,?,?)", 
							sheet.getCell(0, i).getContents(), sheet.getCell(1, i).getContents(),sheet.getCell(2, i).getContents());
			}
		}finally {
			if(wb != null)
				wb.close();
		}
	}
}
