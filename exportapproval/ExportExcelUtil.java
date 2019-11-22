package com.yonyou.yuncai.angang.project.web.controller.exportapproval;
import com.yonyou.yuncai.angang.project.entity.bidproject.CpuBidSectionVO;
import com.yonyou.yuncai.angang.project.entity.bidproject.enums.CpuBidSectionExcleVO;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
/**
 * 这是一个通用的方法，利用了JAVA的反射机制，可以将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上
 *
 * @param title
 *            表格标题名
 * @param headers
 *            表格属性列名数组
 * @param dataset
 *            需要显示的数据集合,集合中一定要放置符合javaBean风格的类的对象。此方法支持的
 *            javaBean属性的数据类型有基本数据类型及String,Date,byte[](图片数据)
 * @param out
 *            与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
 * @param pattern
 *            如果有时间数据，设定输出格式。默认为"yyyy-MM-dd"
 */
public class ExportExcelUtil<T> {
	public void exportExcelStream(OutputStream out, Set set, Set set1, CpuBidSectionVO cpuBidSectionVO, CpuBidSectionExcleVO cpuBidSectionExcleVO, String[] headers) {
		exportExcelStream("湖南黑金时代股份有限公司谈判采购审批单", out, set, set1,cpuBidSectionVO,cpuBidSectionExcleVO,headers);
	}

	public void exportExcelStream(Map<String, Object> map, String[] headers, Collection<T> dataset, OutputStream out,
								  String pattern, CpuBidSectionExcleVO cpuBidSectionExcleVO) {
		exportExcelStream("湖南黑金时代股份有限公司物资招（议）标文件审批单", map, headers, dataset, out, pattern, cpuBidSectionExcleVO);
	}

	@SuppressWarnings({"deprecation"})
	public void exportExcelStream(String title, Map<String, Object> map, String[] headers, Collection<T> dataset,
								  OutputStream out, String pattern, CpuBidSectionExcleVO cpuBidSectionExcleVO) {
		// 声明一个工作薄
		HSSFWorkbook workbook = new HSSFWorkbook();
		// 生成一个表格
		HSSFSheet sheet = workbook.createSheet(title);
		//设置审批流内容表体样式
		//第一行
		HSSFCellStyle boderStyleapprove0 = workbook.createCellStyle();
		boderStyleapprove0.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		HSSFCellStyle boderStyleapprove1 = workbook.createCellStyle();
		boderStyleapprove1.setBorderRight(HSSFCellStyle.BORDER_THIN);

		//第二行
		HSSFCellStyle boderStyleapprove2 = workbook.createCellStyle();
		boderStyleapprove2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		boderStyleapprove2.setAlignment(HSSFCellStyle.ALIGN_CENTER); //居中
		HSSFCellStyle boderStyleapprove2t = workbook.createCellStyle();

		boderStyleapprove2t.setBorderRight(HSSFCellStyle.BORDER_THIN);
		//第三行
		HSSFCellStyle boderStyleapprove3 = workbook.createCellStyle();
		boderStyleapprove3.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		boderStyleapprove3.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		HSSFCellStyle boderStyleapprove3t = workbook.createCellStyle();
		boderStyleapprove3t.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		boderStyleapprove3t.setBorderRight(HSSFCellStyle.BORDER_THIN);

		//设置内容表体样式
		HSSFCellStyle boderStyle = workbook.createCellStyle();
		//垂直居中
		boderStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
		boderStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		boderStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		boderStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		boderStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		//设置合并单元格居中
		//设置单元格文字
		HSSFFont hfont = workbook.createFont();
		hfont.setFontHeightInPoints((short) 13);
		hfont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		boderStyle.setFont(hfont);
		//设置标题样式
		HSSFCellStyle boderStyle1 = workbook.createCellStyle();
		HSSFFont hfont1 = workbook.createFont();
		hfont1.setFontHeightInPoints((short) 13);
		hfont1.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		boderStyle1.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
		boderStyle1.setFont(hfont1);
		//创建合并单元格
		CellRangeAddress cellRangeAddress1 = new CellRangeAddress(0, 1, 0, 1);
		sheet.addMergedRegion(cellRangeAddress1);
		HSSFRow row0 = sheet.createRow(0);
		HSSFRichTextString richStr = new HSSFRichTextString("湖南黑金时代股份有限公司物资招（议）标文件审批单");
		HSSFCell cell0 = row0.createCell(0);
		richStr.applyFont(hfont);
		cell0.setCellValue(richStr);
		cell0.setCellStyle(boderStyle1);
		// 设置表格默认列宽度为35个字节
		sheet.setDefaultColumnWidth((short) 35);
		int j = 0;
		Iterator<T> it = dataset.iterator();
		while (it.hasNext()) {
			T t = (T) it.next();
			// 利用反射，根据javabean属性的先后顺序，动态调用getXxx()方法得到属性值
			Field[] fields = t.getClass().getDeclaredFields();
			for (short i = 0; i < fields.length; i++) {
				Field field = fields[i];
				String fieldName = field.getName();
				//装动态数据
				if (fieldName.equals("listProcurementcomments")) {
					//属性名过滤,包含以下名称的属性名 不会输出
					// 当前属性存在
					if (false) {
						System.out.println(fieldName + "存在");
						continue;
					} else {
						try {
							List<Map<String, String>> listProcurementcomments = cpuBidSectionExcleVO.getListProcurementcomments();
							//审批流遍历索引
							int k = listProcurementcomments.size() - 1;
							for (int index = 0; index < listProcurementcomments.size() * 3; index += 3) {
								//审批人名称
								String procurementcommentspro = listProcurementcomments.get(k).get("Procurementcommentspro");
								//审批人意见
								String procurementcomments = listProcurementcomments.get(k).get("Procurementcomments");
								//提交人名称
								String Opinionsofhandler = listProcurementcomments.get(k).get("Opinionsofhandler");
								//审批时间
								String Data = listProcurementcomments.get(k).get("approvaldata");
								if (k == listProcurementcomments.size() - 1) {
									CellRangeAddress cellRangeAddress = new CellRangeAddress(i + index + 2, i + index + 2, 0, 1);
									sheet.addMergedRegion(cellRangeAddress);
									//第一行
									HSSFRichTextString richStr1 = new HSSFRichTextString(headers[j++]);
									richStr1.applyFont(hfont);
									HSSFRow row = sheet.createRow(i + index + 2);
									HSSFCell cell1 = row.createCell(0);
									cell1.setCellValue(richStr1);
									cell1.setCellStyle(boderStyleapprove0);
									HSSFCell cell = row.createCell(1);
									cell.setCellStyle(boderStyleapprove1);
									//第二行
									CellRangeAddress cellRangeAddress2 = new CellRangeAddress(i + index + 3, i + index + 3, 0, 1);
									sheet.addMergedRegion(cellRangeAddress2);
									HSSFRow row1 = sheet.createRow(i + index + 3);
									HSSFCell cell2 = row1.createCell(0);
									cell2.setCellStyle(boderStyleapprove2);
									HSSFCell cell2t = row1.createCell(1);
									cell2t.setCellStyle(boderStyleapprove2t);
									//第三行
									HSSFRow row3 = sheet.createRow(i + index + 4);
									HSSFCell cell3 = row3.createCell(0);
									cell3.setCellValue("姓名:" + Opinionsofhandler);
									cell3.setCellStyle(boderStyleapprove3);
									HSSFCell cell3t = row3.createCell(1);
									cell3t.setCellValue(Data);
									cell3t.setCellStyle(boderStyleapprove3t);
									k--;
								} else {
									CellRangeAddress cellRangeAddress = new CellRangeAddress(i + index + 2, i + index + 2, 0, 1);
									sheet.addMergedRegion(cellRangeAddress);
									//第一行
									HSSFRichTextString richStr1 = new HSSFRichTextString(headers[j++]);
									richStr1.applyFont(hfont);
									HSSFRow row = sheet.createRow(i + index + 2);
									HSSFCell cell1 = row.createCell(0);
									cell1.setCellValue(richStr1);
									cell1.setCellStyle(boderStyleapprove0);
									HSSFCell cell = row.createCell(1);
									cell.setCellStyle(boderStyleapprove1);
									//第二行
									CellRangeAddress cellRangeAddress2 = new CellRangeAddress(i + index + 3, i + index + 3, 0, 1);
									sheet.addMergedRegion(cellRangeAddress2);
									HSSFRow row1 = sheet.createRow(i + index + 3);
									HSSFCell cell2 = row1.createCell(0);
									cell2.setCellValue(procurementcomments);
									cell2.setCellStyle(boderStyleapprove2);
									HSSFCell cell2t = row1.createCell(1);
									cell2t.setCellStyle(boderStyleapprove2t);
									//第三行
									HSSFRow row3 = sheet.createRow(i + index + 4);
									HSSFCell cell3 = row3.createCell(0);
									cell3.setCellValue("姓名:" + procurementcommentspro);
									cell3.setCellStyle(boderStyleapprove3);
									HSSFCell cell3t = row3.createCell(1);
									cell3t.setCellValue(Data);
									cell3t.setCellStyle(boderStyleapprove3t);
									k--;
								}
							}
						} catch (SecurityException e) {
							e.printStackTrace();
						} finally {
							// 清理资源
						}
					}
				}//装静态数据
				else {
					HSSFRichTextString richStr1 = new HSSFRichTextString(headers[j++]);
					richStr1.applyFont(hfont);
					HSSFRow row = sheet.createRow(i + 2);
					HSSFCell cell1 = row.createCell(0);
					cell1.setCellValue(richStr1);
					cell1.setCellStyle(boderStyle);
					//属性名过滤,包含以下名称的属性名 不会输出
					// 当前属性存在
					if (false) {
						System.out.println(fieldName + "存在");
						continue;
					} else {
						System.out.println("可打印属性" + fieldName);
						HSSFCell cell = row.createCell(1);
						cell.setCellStyle(boderStyle);
						String getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
						try {
							Class<?> tCls = t.getClass();
							Method getMethod = tCls.getMethod(getMethodName, new Class[]{});
							Object value = getMethod.invoke(t, new Object[]{});
							// 判断值的类型后进行强制类型转换
							String textValue = null;
							if (value instanceof Boolean) {
								boolean bValue = (Boolean) value;
								textValue = "1";
								if (!bValue) {
									textValue = "0";
								}
							} else if (value instanceof Date) {
								Date date = (Date) value;
								SimpleDateFormat sdf = new SimpleDateFormat(pattern);
								textValue = sdf.format(date);
							} else {
								// 其它数据类型都当作字符串简单处理
								if (value == null) {
									value = "";
								}
								textValue = value.toString();
							}
							// 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
							if (textValue != null) {
								Pattern p = Pattern.compile("^//d+(//.//d+)?$");
								Matcher matcher = p.matcher(textValue);
								if (matcher.matches()) {
									// 是数字当作double处理Double.parseDouble(textValue) 会让你自动换算
									HSSFRichTextString richString = new HSSFRichTextString(textValue);
									HSSFFont font3 = workbook.createFont();
									richString.applyFont(font3);
									cell.setCellValue(richString);
								} else {
									HSSFRichTextString richString = new HSSFRichTextString(textValue);
									HSSFFont font3 = workbook.createFont();
//							font3.setColor(HSSFColor.BLUE.index);
									richString.applyFont(font3);
									cell.setCellValue(richString);
								}
							}
						} catch (SecurityException e) {
							e.printStackTrace();
						} catch (NoSuchMethodException e) {
							e.printStackTrace();
						} catch (IllegalArgumentException e) {
							e.printStackTrace();
						} catch (IllegalAccessException e) {
							e.printStackTrace();
						} catch (InvocationTargetException e) {
							e.printStackTrace();
						} finally {
							// 清理资源
						}
					}
				}
			}
		}
		workbook.getSheetAt(0).getWorkbook().getAllNames();
		try {
			workbook.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	@SuppressWarnings({"deprecation"})
	public void exportExcelStream(String title,OutputStream out, Set set, Set set1,
								  CpuBidSectionVO cpuBidSectionVO, CpuBidSectionExcleVO cpuBidSectionExcleVO,
								  String[] headers) {
        // 声明一个工作薄
		HSSFWorkbook workbook = new HSSFWorkbook();
		// 生成一个表格
		HSSFSheet sheet = workbook.createSheet(title);
		//设置审批流内容表体样式
		//设置单元格文字
		HSSFFont hfont = workbook.createFont();
		hfont.setFontHeightInPoints((short) 13);
		hfont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		//第一行
		HSSFCellStyle approve00 = workbook.createCellStyle();
		approve00.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		approve00.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
		approve00.setFont(hfont);

		HSSFCellStyle approve0 = workbook.createCellStyle();
		approve0.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
		approve0.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		approve0.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		approve0.setFont(hfont);

		HSSFCellStyle approve1 = workbook.createCellStyle();
		approve1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		approve1.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
		approve1.setBorderRight(HSSFCellStyle.BORDER_THIN);
		approve1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		approve1.setFont(hfont);
        //审批流第一行
		HSSFCellStyle boderStyleapprove0 = workbook.createCellStyle();
		boderStyleapprove0.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		HSSFCellStyle boderStyleapprove1 = workbook.createCellStyle();
		boderStyleapprove1.setBorderRight(HSSFCellStyle.BORDER_THIN);
		//第二行
		HSSFCellStyle boderStyleapprove2 = workbook.createCellStyle();
		boderStyleapprove2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		boderStyleapprove2.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
		HSSFCellStyle boderStyleapprove2t = workbook.createCellStyle();
		boderStyleapprove2t.setBorderRight(HSSFCellStyle.BORDER_THIN);
		//第三行
		HSSFCellStyle boderStyleapprove3 = workbook.createCellStyle();
		boderStyleapprove3.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		boderStyleapprove3.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		HSSFCellStyle boderStyleapprove3t = workbook.createCellStyle();
		boderStyleapprove3t.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		boderStyleapprove3t.setBorderRight(HSSFCellStyle.BORDER_THIN);
		//设置内容表体样式
		HSSFCellStyle boderStyle = workbook.createCellStyle();
		//垂直居中
		boderStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
		boderStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		boderStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		boderStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		boderStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		//设置合并单元格居中
		boderStyle.setFont(hfont);
		//设置标题样式
		HSSFCellStyle boderStyle1 = workbook.createCellStyle();
		HSSFFont hfont1 = workbook.createFont();
		hfont1.setFontHeightInPoints((short) 13);
		hfont1.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		boderStyle1.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
		boderStyle1.setFont(hfont1);
		//创建合并单元格
		CellRangeAddress cellRangeAddress1 = new CellRangeAddress(0, 1, 0, 1);
		sheet.addMergedRegion(cellRangeAddress1);
		HSSFRow row0 = sheet.createRow(0);
		HSSFRichTextString richStr = new HSSFRichTextString("湖南黑金时代股份有限公司谈判采购审批单");
		HSSFCell cell0 = row0.createCell(0);
		richStr.applyFont(hfont);
		cell0.setCellValue(richStr);
		cell0.setCellStyle(boderStyle1);
		// 设置表格默认列宽度为35个字节
		sheet.setDefaultColumnWidth((short) 35);
		//招（议）标文件编号
		HSSFRow row1 = sheet.createRow(2);
		HSSFCell cell1 = row1.createCell(0);
		cell1.setCellValue("招（议）标文件编号");
		cell1.setCellStyle(boderStyle);
		HSSFCell cell1t = row1.createCell(1);
		cell1t.setCellValue(cpuBidSectionVO.getBidSectionCode());
		cell1t.setCellStyle(boderStyle);
		//招（议）标项目
		HSSFRow row2 = sheet.createRow(3);
		HSSFCell cell2 = row2.createCell(0);
		cell2.setCellValue("招（议）标项目");
		cell2.setCellStyle(boderStyle);
		HSSFCell cell2t = row2.createCell(1);
		cell2t.setCellValue(cpuBidSectionVO.getBidSectionName());
		cell2t.setCellStyle(boderStyle);
		//物料
		int size = set1.size();
		Iterator it =set1.iterator();
		int i=0;
		while (it.hasNext()){
			HSSFRow row3 = sheet.createRow(4+i);
			HSSFCell cell3 = row3.createCell(0);
			if(i==0){
				cell3.setCellValue("谈判采购物资名称");
				cell3.setCellStyle(approve00);
			}else if(i==size-1){
				cell3.setCellStyle(approve0);
			}
			else{
				cell3.setCellStyle(approve00);
			}
			HSSFCell cell3t = row3.createCell(1);
			cell3t.setCellValue((String) it.next());
			cell3t.setCellStyle(approve1);
			i++;
		}
		//供应商
		int size1 = set.size();
		Iterator it1 =set.iterator();
		int j=0;
		while (it1.hasNext()){
			HSSFRow row3 = sheet.createRow(4+j+size);
			HSSFCell cell3 = row3.createCell(0);
			if(j==0){
				cell3.setCellValue("谈判采购供应商名称");
				cell3.setCellStyle(approve00);
			}else if(j==size1-1){
				cell3.setCellStyle(approve0);
			}
			else{
				cell3.setCellStyle(approve00);
			}
			HSSFCell cell3t = row3.createCell(1);
			cell3t.setCellValue((String) it1.next());
			cell3t.setCellStyle(approve1);
			j++;
		}
		//拼审批流单元格的起始位置
		int one = 2+j+size;
		//-----------------------审批流显示--------------------------------------------------
		List<Map<String, String>> listProcurementcomments = cpuBidSectionExcleVO.getListProcurementcomments();
		//审批流遍历索引
		int k = listProcurementcomments.size() - 1;
		for (int index = 0; index < listProcurementcomments.size() * 3; index += 3) {
			//审批人名称
			String procurementcommentspro = listProcurementcomments.get(k).get("Procurementcommentspro");
			//审批人意见
			String procurementcomments = listProcurementcomments.get(k).get("Procurementcomments");
			//提交人名称
			String Opinionsofhandler = listProcurementcomments.get(k).get("Opinionsofhandler");
			//审批时间
			String Data = listProcurementcomments.get(k).get("approvaldata");
			if (k == listProcurementcomments.size() - 1) {
				CellRangeAddress cellRangeAddress = new CellRangeAddress(one + index + 2, one + index + 2, 0, 1);
				sheet.addMergedRegion(cellRangeAddress);
				//第一行
				HSSFRichTextString richStr1 = new HSSFRichTextString(headers[j++]);
				richStr1.applyFont(hfont);
				HSSFRow row4 = sheet.createRow(one + index + 2);
				HSSFCell cell4 = row4.createCell(0);
				cell4.setCellValue(richStr1);
				cell4.setCellStyle(boderStyleapprove0);
				HSSFCell cell4t = row4.createCell(1);
				cell4t.setCellStyle(boderStyleapprove1);
				//第二行
				CellRangeAddress cellRangeAddress2 = new CellRangeAddress(one + index + 3, one + index + 3, 0, 1);
				sheet.addMergedRegion(cellRangeAddress2);
				HSSFRow row5 = sheet.createRow(one + index + 3);
				HSSFCell cell5 = row5.createCell(0);
				cell5.setCellStyle(boderStyleapprove2);
				HSSFCell cell5t = row5.createCell(1);
				cell5t.setCellStyle(boderStyleapprove2t);
				//第三行
				HSSFRow row6 = sheet.createRow(one + index + 4);
				HSSFCell cell6 = row6.createCell(0);
				cell6.setCellValue("姓名:" + Opinionsofhandler);
				cell6.setCellStyle(boderStyleapprove3);
				HSSFCell cell6t = row6.createCell(1);
				cell6t.setCellValue(Data);
				cell6t.setCellStyle(boderStyleapprove3t);
				k--;
			} else {
				CellRangeAddress cellRangeAddress = new CellRangeAddress(one + index + 2, one + index + 2, 0, 1);
				sheet.addMergedRegion(cellRangeAddress);
				//第一行
				HSSFRichTextString richStr1 = new HSSFRichTextString(headers[j++]);
				richStr1.applyFont(hfont);
				HSSFRow row4 = sheet.createRow(one + index + 2);
				HSSFCell cell4 = row4.createCell(0);
				cell4.setCellValue(richStr1);
				cell4.setCellStyle(boderStyleapprove0);
				HSSFCell cell4t = row4.createCell(1);
				cell4t.setCellStyle(boderStyleapprove1);
				//第二行
				CellRangeAddress cellRangeAddress2 = new CellRangeAddress(one + index + 3, one + index + 3, 0, 1);
				sheet.addMergedRegion(cellRangeAddress2);
				HSSFRow row5 = sheet.createRow(one + index + 3);
				HSSFCell cell5 = row5.createCell(0);
				cell5.setCellValue(procurementcomments);
				cell5.setCellStyle(boderStyleapprove2);
				HSSFCell cell5t = row5.createCell(1);
				cell5t.setCellStyle(boderStyleapprove2t);
				//第三行
				HSSFRow row6 = sheet.createRow(one + index + 4);
				HSSFCell cell6 = row6.createCell(0);
				cell6.setCellValue("姓名:" + procurementcommentspro);
				cell6.setCellStyle(boderStyleapprove3);
				HSSFCell cell6t = row6.createCell(1);
				cell6t.setCellValue(Data);
				cell6t.setCellStyle(boderStyleapprove3t);
				k--;
			}
		}
		workbook.getSheetAt(0).getWorkbook().getAllNames();
		try {
			workbook.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}