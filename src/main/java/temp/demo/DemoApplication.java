package temp.demo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;

import com.drew.imaging.ImageMetadataReader;
import com.drew.metadata.Directory;
import com.drew.metadata.Metadata;
import com.drew.metadata.Tag;
import com.drew.metadata.exif.ExifIFD0Directory;

import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Scanner;
import java.util.concurrent.Executors;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SpringBootApplication
public class DemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);
		List<String> temp = new ArrayList<String>();
		String allFilePath;
		String tifFolderPath;
		// temp = getInput();
		// allFilePath = temp.get(0);
		allFilePath = "D:\\Program Files\\AMSKY\\TiffDownload (x64) Ver4.1.3 R\\Log\\TiffDownloadLog\\Aurora3LV256_800T(83080000)\\";
		Date date = new Date();
		SimpleDateFormat dateFormat= new SimpleDateFormat("yyyy-MM-dd");
		allFilePath = allFilePath + dateFormat.format(date) + ".ALL";
		// tifFolderPath = temp.get(1);
		tifFolderPath = "E:\\";
		try {
			// 获取tifList
			List<String> tifList = new ArrayList<String>();
			BufferedReader reader = new BufferedReader(new FileReader(allFilePath));
			String line;
			while ((line = reader.readLine()) != null) {
				String pattern = "作业 (.*?) 打印完成！";
				Pattern regex = Pattern.compile(pattern);
        		Matcher matcher = regex.matcher(line);
				if (matcher.find()) {
					String extractedContent = matcher.group(1);
					tifList.add(extractedContent);
				}
			}

			String[] folderList = {"760x605", "785x580", "890x633", "890x655", "904x605", "914x661", "978x655", "1030x681", "1030x800"};
			// String templeteName;
			AtomicReference<String> templeteName = new AtomicReference<>();
			String[][] data = new String[tifList.size()][2];
			ThreadPoolExecutor threadPoolExecutor = new ThreadPoolExecutor(
			4,
			4,
			0L,
			TimeUnit.MILLISECONDS,
			new LinkedBlockingQueue<>(),
			Executors.defaultThreadFactory(),
			new ThreadPoolExecutor.AbortPolicy());
			threadPoolExecutor.execute(() -> {
				for (int i = 0; i < tifList.size(); i++) {
					for (int j = 0; j < folderList.length; j++) {
						templeteName.set(searchFile(tifFolderPath + folderList[j] + "\\OkFiles\\", tifList.get(i)));
						if (!(templeteName.get() == null)) {
							data[i][0] = folderList[j];
							data[i][1] = tifList.get(i);
							break;
						}
					}
				}
			});
			threadPoolExecutor.shutdown();
			threadPoolExecutor.awaitTermination(Long.MAX_VALUE, TimeUnit.NANOSECONDS);

			// // 获得模板宽高
			// String[][] data = new String[tifList.size()][2];
			// for (int i = 0; i < tifList.size(); i++) {
			// 	String tiffPath = tifFolderPath + tifList.get(i);
        	// 	File tiffFile = new File(tiffPath);
			// 	Metadata metadata = ImageMetadataReader.readMetadata(tiffFile);
			// 	double xResolution = 0.0;
			// 	double yResolution = 0.0;
			// 	// 遍历所有目录（Directory）
			// 	for (Directory directory : metadata.getDirectories()) {
			// 		// 遍历目录中的标签（Tag）
			// 		for (Tag tag : directory.getTags()) {
			// 			if (tag.getTagName().equals("X Resolution")) {
			// 				xResolution = Integer.parseInt(tag.getDescription().split(" ")[0]);
			// 			} else if (tag.getTagName().equals("Y Resolution")) {
			// 				yResolution = Integer.parseInt(tag.getDescription().split(" ")[0]);
			// 			}
			// 		}
			// 	}
			// 	// 获取图像的宽度和高度
			// 	double imageWidth = metadata.getFirstDirectoryOfType(ExifIFD0Directory.class).getInt(ExifIFD0Directory.TAG_IMAGE_WIDTH);
			// 	double imageHeight = metadata.getFirstDirectoryOfType(ExifIFD0Directory.class).getInt(ExifIFD0Directory.TAG_IMAGE_HEIGHT);
			// 	int templeteWidth = (int)Math.round(imageWidth / xResolution * 25.4);
			// 	int templeteHeight = (int)Math.round(imageHeight / yResolution * 25.4);
			// 	String templeteName = Integer.toString(templeteHeight) + 'x' + Integer.toString(templeteWidth);
			// 	data[i][0] = templeteName;
			// 	data[i][1] = tifList.get(i);
			// }

			// 创建excel
			Workbook workbook;
			Sheet sheet;
			File folder = new File("E:\\统计");
			SimpleDateFormat dateFormatExcelFile= new SimpleDateFormat("yyyy-MM");
			int rowIndex = 1;
			if(!folder.exists()) {
				folder.mkdirs();
			}
			File file = new File("E:\\统计\\" + dateFormatExcelFile.format(date) + ".xlsx");
			boolean appendToFile = file.exists();
			if (!appendToFile) {
				workbook = new XSSFWorkbook();
				sheet = workbook.createSheet("Sheet1");
				// 创建表头
				Row headerRow = sheet.createRow(0);
				String[] headers = {"模板名", "作业名"};
				for (int i = 0; i < headers.length; i++) {
					Cell cell = headerRow.createCell(i);
					cell.setCellValue(headers[i]);
				}
			} else {
				workbook = WorkbookFactory.create(file);
				sheet = workbook.getSheetAt(0); // 假设工作簿中的第一个表格
				int lastRowIndex = sheet.getLastRowNum(); // 获取最后一行的索引
				rowIndex = lastRowIndex + 1;
				file.delete();
			}
			// 插入表数据
			for (String[] rowData : data) {
				Row row = sheet.createRow(rowIndex++);
				int columnIndex = 0;
				for (String value : rowData) {
					Cell cell = row.createCell(columnIndex++);
					cell.setCellValue(value);
				}
			}
			FileOutputStream outputStream = new FileOutputStream(file);
			// FileOutputStream outputStream = new FileOutputStream(file, appendToFile);
			workbook.write(outputStream);
            System.out.println("Excel文件创建成功！");
			outputStream.close();
		} catch (Exception e) {
			// TODO: handle exception
			System.out.println("=====================" + e);
		}
	}

	private static List<String> getInput() {
		Scanner scanner = new Scanner(System.in);
        System.out.print("请输入ALL文件路径: ");
        String allFilePath = scanner.nextLine();
        System.out.println("您输入的ALL文件路径是：" + allFilePath);
        System.out.print("请输入tif文件夹路径: ");
        String tifFolderPath = scanner.nextLine();
        System.out.println("您输入的tif文件夹路径是：" + tifFolderPath);
		List<String> arr = new ArrayList<String>();
		arr.add(allFilePath);
		arr.add(tifFolderPath);
		return arr;
	}

	private static String searchFile(String folderPath, String fileName) {
        File folder = new File(folderPath);
        if (!folder.exists() || !folder.isDirectory()) {
            return null;
        }
        File[] files = folder.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isFile() && file.getName().equals(fileName)) {
                    return file.getAbsolutePath();
                }
            }
        }
        return null;
    }

}

// public class DemoApplication implements CommandLineRunner {

// 	public static void main(String[] args) {
// 		SpringApplication.run(DemoApplication.class, args);
// 	}

// 	@Override
//     public void run(String... args) {
//         Scanner scanner = new Scanner(System.in);
//         System.out.print("请输入ALL文件路径: ");
//         String allFilePath = scanner.nextLine();
//         System.out.println("您输入的ALL文件路径是：" + allFilePath);
//         System.out.print("请输入tif文件夹路径: ");
//         String tifFolderPath = scanner.nextLine();
//         System.out.println("您输入的tif文件夹路径是：" + tifFolderPath);

// 		try {
// 			// 获取tifList
// 			List<String> tifList = new ArrayList<String>();
// 			BufferedReader reader = new BufferedReader(new FileReader(allFilePath));
// 			String line;
// 			while ((line = reader.readLine()) != null) {
// 				String pattern = "作业 (.*?) 打印完成！";
// 				Pattern regex = Pattern.compile(pattern);
//         		Matcher matcher = regex.matcher(line);
// 				if (matcher.find()) {
// 					String extractedContent = matcher.group(1);
// 					tifList.add(extractedContent);
// 				}
// 			}

// 			// 获得模板宽高
// 			String[][] data = new String[tifList.size()][2];
// 			for (int i = 0; i < tifList.size(); i++) {
// 				String tiffPath = tifFolderPath + tifList.get(i);
//         		File tiffFile = new File(tiffPath);
// 				Metadata metadata = ImageMetadataReader.readMetadata(tiffFile);
// 				double xResolution = 0.0;
// 				double yResolution = 0.0;
// 				// 遍历所有目录（Directory）
// 				for (Directory directory : metadata.getDirectories()) {
// 					// 遍历目录中的标签（Tag）
// 					for (Tag tag : directory.getTags()) {
// 						if (tag.getTagName().equals("X Resolution")) {
// 							xResolution = Integer.parseInt(tag.getDescription().split(" ")[0]);
// 						} else if (tag.getTagName().equals("Y Resolution")) {
// 							yResolution = Integer.parseInt(tag.getDescription().split(" ")[0]);
// 						}
// 					}
// 				}
// 				// 获取图像的宽度和高度
// 				double imageWidth = metadata.getFirstDirectoryOfType(ExifIFD0Directory.class).getInt(ExifIFD0Directory.TAG_IMAGE_WIDTH);
// 				double imageHeight = metadata.getFirstDirectoryOfType(ExifIFD0Directory.class).getInt(ExifIFD0Directory.TAG_IMAGE_HEIGHT);
// 				int templeteWidth = (int)Math.round(imageWidth / xResolution * 25.4);
// 				int templeteHeight = (int)Math.round(imageHeight / yResolution * 25.4);
// 				String templeteName = Integer.toString(templeteHeight) + 'x' + Integer.toString(templeteWidth);
// 				data[i][0] = templeteName;
// 				data[i][1] = tifList.get(i);
// 			}

// 			// 创建excel
// 			Workbook workbook = new XSSFWorkbook();
// 			Sheet sheet = workbook.createSheet("Sheet1");
// 			// 创建表头
// 			Row headerRow = sheet.createRow(0);
// 			String[] headers = {"模板名", "作业名"};
// 			for (int i = 0; i < headers.length; i++) {
// 				Cell cell = headerRow.createCell(i);
// 				cell.setCellValue(headers[i]);
// 			}
// 			// 插入表数据
// 			int rowIndex = 1;
// 			for (String[] rowData : data) {
// 				Row row = sheet.createRow(rowIndex++);
// 				int columnIndex = 0;
// 				for (String value : rowData) {
// 					Cell cell = row.createCell(columnIndex++);
// 					cell.setCellValue(value);
// 				}
// 			}
// 			FileOutputStream outputStream = new FileOutputStream(tifFolderPath + "example.xlsx");
// 			workbook.write(outputStream);
//             System.out.println("Excel文件创建成功！");
// 		} catch (Exception e) {
// 			// TODO: handle exception
// 			System.out.println("=====================" + e);
// 		}
//         scanner.close();
//     }

// }
