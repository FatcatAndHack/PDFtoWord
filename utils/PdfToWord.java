package com.ducway.framework.modular.zmxzApplication.controller.utils;

import com.spire.pdf.FileFormat;
import com.spire.pdf.PdfDocument;
import com.spire.pdf.widget.PdfPageCollection;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.swing.filechooser.FileSystemView;
import java.io.*;
import java.util.List;


public class PdfToWord extends PdfDocument {

	// 1、如果是大文件，需要进行切分，保存的子pdf路径
	String splitPath = "./split/";

	// 2、如果是大文件，需要对子pdf文件一个一个进行转化
	String docPath = "./doc/";


	public  String pdftoword(String  srcPath , String fileName) {
		// 3、最终生成的doc所在的目录，默认是和引入的一个地方，开源时对外提供下载的接口。
		String desPath = srcPath.substring(0, srcPath.length()-4)+".docx";
		boolean result = false;
		try {
			// 0、判断输入的是否是pdf文件
			//第一步：判断输入的是否合法
			boolean flag = isPDFFile(srcPath);
			//第二步：在输入的路径下新建文件夹
			boolean flag1 = create();

			if (flag && flag1) {
				// 1、加载pdf
				PdfDocument pdf = new PdfDocument();
				pdf.loadFromFile(srcPath);
				PdfPageCollection num = pdf.getPages();
				PdfDocument sonpdf = new PdfDocument();

				fileName = fileName.split("\\.")[0];
				// 2、如果pdf的页数小于11，那么直接进行转化
				if (num.getCount() <= 10) {
					sonpdf.loadFromFile(srcPath);

					sonpdf.saveToFile(docPath+"test.docx",FileFormat.DOCX);
					InputStream is = new FileInputStream(docPath+"test.docx");
					XWPFDocument document = new XWPFDocument(is);
					document.removeBodyElement(0);
					FileSystemView fsv = FileSystemView.getFileSystemView();
					File path=fsv.getHomeDirectory();
					OutputStream os=new FileOutputStream(path+"\\"+fileName+".docx"); //生成到桌面
					try {
						document.write(os);
					} catch (Exception e) {
						e.printStackTrace();
					}finally {
						new FileDeleteUtils().clearFiles(splitPath);
						new FileDeleteUtils().clearFiles(docPath);
					}
				}
				// 3、否则输入的页数比较多，就开始进行切分再转化
				else {
					// 第一步：将其进行切分,每页一张pdf
					pdf.split(splitPath+"test{0}.pdf",0);

					// 第二步：将切分的pdf，一个一个进行转换
					File[] fs = getSplitFiles(splitPath);
					for(int i=0;i<fs.length;i++) {
						sonpdf.loadFromFile(fs[i].getAbsolutePath());
						//删除第一个水印
						sonpdf.saveToFile(docPath+fs[i].getName().substring(0, fs[i].getName().length()-4)+".docx",FileFormat.DOCX);
						InputStream is = new FileInputStream(docPath+fs[i].getName().substring(0, fs[i].getName().length()-4)+".docx");
						XWPFDocument document = new XWPFDocument(is);
						//以上Spire.Doc 生成的文件会自带警告信息，这里来删除Spire.Doc 的警告
						List<XWPFParagraph> paragraphs = document.getParagraphs();
						int secIndex = -1;
						for (int index = 0; index < paragraphs.size(); index++) {
							boolean findFlag = false;
							List<XWPFRun> runs = paragraphs.get(index).getRuns();
							for (XWPFRun run : runs) {
								if(     " Evaluation Warning : The document was created with Spire.PDF for Java.".equalsIgnoreCase(run.toString().trim()) ||
										"Evaluation Warning : The document was created with Spire.PDF for Java.".equalsIgnoreCase(run.toString().trim())   ||
										"Evaluation Warning: The document was created with Spire.Doc for JAVA.".equalsIgnoreCase(run.toString().trim())   ||
										"The document was created with Spire.PDF for Java.".equalsIgnoreCase(run.toString().trim())){
									findFlag = true;
									break;
								}
							}
							if(findFlag){
								//返回索引
								secIndex = index;
								document.removeBodyElement(secIndex);
								break;
							}
						}
						//删除原来的 生成新的
						OutputStream os=new FileOutputStream(docPath+fs[i].getName().substring(0, fs[i].getName().length()-4)+".docx");
						try {
							document.write(os);
							clearAgain(docPath+fs[i].getName().substring(0, fs[i].getName().length()-4)+".docx");
						} catch (Exception e) {
							e.printStackTrace();
						}
					}
					//第三步：对转化的doc文档进行合并，合并成一个大的word
					try {
						result = MergeWordDocument.merge(docPath, desPath);
						System.out.println(desPath);
						//获取生成路径 删除第一个
						InputStream is = new FileInputStream(desPath);
						XWPFDocument document = new XWPFDocument(is);
						//以上Spire.Doc 生成的文件会自带警告信息，这里来删除Spire.Doc 的警告
						document.removeBodyElement(0);
						OutputStream os=new FileOutputStream(desPath);
						try {
							document.write(os);
						} catch (Exception e) {
							e.printStackTrace();
						}
					} catch (Exception e) {
						e.printStackTrace();
					}
				}
			} else {
				System.out.println("输入的不是pdf文件");
				return "输入的不是pdf文件";
			}
		} catch (Exception e) {
			e.printStackTrace();
		}finally {
			//4、把刚刚缓存的split和doc删除
			if(result==true) {
				new FileDeleteUtils().clearFiles(splitPath);
				new FileDeleteUtils().clearFiles(docPath);
			}
		}
		return fileName.split("\\.")[0] + ".docx";
	}

	private void deleteFile(File file){
		if(file.isDirectory()){
			File[] files = file.listFiles();
			for(int i=0; i<files.length; i++){
				deleteFile(files[i]);
			}
		}
		file.delete();
	}

	private  boolean create() {
		File f = new File(splitPath);
		File f1 = new File(docPath);
		if(!f.exists() )  f.mkdirs();
		if(!f.exists() )  f1.mkdirs();
		return true;
	}

	// 判断是否是pdf文件
	private  boolean isPDFFile(String srcPath2) {
		File file = new File(srcPath2);
		String filename = file.getName();
		if (filename.endsWith(".pdf")) {
			return true;
		}
		return false;
	}

	// 取得某一路径下所有的pdf
	private  File[] getSplitFiles(String path) {
		File f = new File(path);
		File[] fs = f.listFiles();
		if (fs == null) {
			return null;
		}
		return fs;
	}

	private synchronized void  clearAgain(String docPath) throws IOException {
		InputStream is = new FileInputStream(docPath);
		XWPFDocument document = new XWPFDocument(is);
		//以上Spire.Doc 生成的文件会自带警告信息，这里来删除Spire.Doc 的警告
		List<XWPFParagraph> paragraphs = document.getParagraphs();
		int secIndex = -1;
		for (int i = 0; i < paragraphs.size(); i++) {
			boolean flag = false;
			List<XWPFRun> runs = paragraphs.get(i).getRuns();
			for (XWPFRun run : runs) {
				if(    run.toString().contains("Spire.Doc")	||
						" Evaluation Warning : The document was created with Spire.PDF for Java.".equalsIgnoreCase(run.toString().trim()) ||
						"Evaluation Warning : The document was created with Spire.PDF for Java.".equalsIgnoreCase(run.toString().trim())   ||
						"Evaluation Warning: The document was created with Spire.Doc for JAVA.".equalsIgnoreCase(run.toString().trim())   ||
						" Evaluation Warning: The document was created with Spire.Doc for JAVA.".equalsIgnoreCase(run.toString().trim())   ||
						"The document was created with Spire.PDF for Java.".equalsIgnoreCase(run.toString().trim())){
					flag = true;
					break;
				}
			}
			if(flag){
				//返回索引
				secIndex = i;
				document.removeBodyElement(secIndex);
				break;
			}
		}
		//输出word内容文件流，新输出路径位置
		OutputStream os=new FileOutputStream(docPath);
		try {
			document.write(os);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}