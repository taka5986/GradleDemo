package com.inesat.shcc.service;

import com.inesat.shcc.api.dto.ChainCheckReportDTO;
import com.inesat.shcc.repository.EmployeeVaccRepository;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@Log4j2
public class ExcelService {
    @Autowired
    private EmployeeVaccRepository excelReportRepository;

    public InputStream createExcel(String date) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("从业人员核酸检测及疫苗接种统计表");

        //设置字体
        Font headFont = workbook.createFont();
        headFont.setFontHeightInPoints((short) 12);
        headFont.setFontName("宋体");

        //设置头部单元格样式
        CellStyle headStyle = workbook.createCellStyle();
        headStyle.setBorderBottom(BorderStyle.THIN);  //设置单元格线条
        headStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());   //设置单元格颜色
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        headStyle.setBorderRight(BorderStyle.THIN);
        headStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        headStyle.setBorderTop(BorderStyle.THIN);
        headStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        headStyle.setAlignment(HorizontalAlignment.CENTER);    //设置水平对齐方式
        headStyle.setVerticalAlignment(VerticalAlignment.CENTER);  //设置垂直对齐方式
        headStyle.setWrapText(true);
        headStyle.setFont(headFont);  //设置字体

        /*设置数据单元格格式*/
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setBorderBottom(BorderStyle.THIN);  //设置单元格线条
        dataStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());   //设置单元格颜色
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setAlignment(HorizontalAlignment.LEFT);    //设置水平对齐方式
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);  //设置垂直对齐方式
        dataStyle.setWrapText(true);
        dataStyle.setFont(headFont);  //设置字体

        // 标题
        HSSFRow row0 = sheet.createRow(0);
        HSSFCell cell00 = row0.createCell(0);
        cell00.setCellStyle(headStyle);
        cell00.setCellValue("从业人员核酸检测及疫苗接种统计表");
        CellRangeAddress cra0 =new CellRangeAddress(0, 0, 0, 19);

        //统计日期
        HSSFRow row1 = sheet.createRow(1);
        HSSFCell cell10 = row1.createCell(0);
        cell10.setCellStyle(dataStyle);
        cell10.setCellValue("统计日期：" + date);
        CellRangeAddress craDate = new CellRangeAddress(1, 1, 0, 19);
        sheet.addMergedRegion(craDate);
        setBorder(craDate,sheet);

        HSSFRow row2 = sheet.createRow(2);
        String columns1[] = {"区域", "冷库名称", "库类", "冷链从业人员总人数（在岗）", "核酸检测人数"};
        for (int i = 0; i < columns1.length; i++) {
            HSSFCell cell = row2.createCell(i);
            cell.setCellValue(columns1[i]);
            cell.setCellStyle(headStyle);
        }
        HSSFCell cell16 = row2.createCell(8);
        cell16.setCellValue("疫苗接种人数");
        cell16.setCellStyle(headStyle);

        HSSFCell cell18 = row2.createCell(17);
        cell18.setCellValue("加强针接种人数");
        cell18.setCellStyle(headStyle);

        HSSFRow row3 = sheet.createRow(3);
        String columns2[] = {"今日", "累计", "检测率","超期", "应接种人数","今日(仅一剂)第一剂次", "累计(仅一剂)第一剂次",
                "今日(二剂)第一剂次", "今日(二剂)第二剂次", "累计(二剂)第一剂次", "累计(二剂)第二剂次", "(二剂)第二次完成率", "接种完成率",
                "应接种人数","已接种人数","接种完成率"};
        for (int i = 0; i < columns2.length; i++) {
            HSSFCell cell = row3.createCell(i + 4);
            cell.setCellValue(columns2[i]);
            cell.setCellStyle(headStyle);
        }

        CellRangeAddress cra1 = new CellRangeAddress(2, 3, 0, 0);
        CellRangeAddress cra2 = new CellRangeAddress(2, 3, 1, 1);
        CellRangeAddress cra3 = new CellRangeAddress(2, 3, 2, 2);
        CellRangeAddress cra4 = new CellRangeAddress(2, 3, 3, 3);
        CellRangeAddress cra5 = new CellRangeAddress(2, 2, 4, 7);
        CellRangeAddress cra6 = new CellRangeAddress(2, 2, 8, 16);
        CellRangeAddress cra7 = new CellRangeAddress(2, 2, 17, 19);

        // 设置合并单元格边框
        sheet.addMergedRegion(cra0);
        sheet.addMergedRegion(cra1);
        sheet.addMergedRegion(cra2);
        sheet.addMergedRegion(cra3);
        sheet.addMergedRegion(cra4);
        sheet.addMergedRegion(cra5);
        sheet.addMergedRegion(cra6);
        sheet.addMergedRegion(cra7);

        setBorder(cra0,sheet);
        setBorder(cra1,sheet);
        setBorder(cra2,sheet);
        setBorder(cra3,sheet);
        setBorder(cra4,sheet);
        setBorder(cra5,sheet);
        setBorder(cra6,sheet);
        setBorder(cra7,sheet);
        /*
         * 数据项仅含 有在职状态从业人员得冷库数据
         * 实际查询时间为（传入日期 - 6：00 --- 传入日期 + 18：00）
         */
        //疫苗接种情况统计
        List<ChainCheckReportDTO> lt1 = excelReportRepository.getVaccTotalData(date);
        //核酸检测统计
        List<ChainCheckReportDTO> lt2 = excelReportRepository.getHSData(date);
        Map<String, ChainCheckReportDTO> map = new HashMap<>();
        if(!CollectionUtils.isEmpty(lt2)){
            map = lt2.stream().collect(Collectors.toMap(ChainCheckReportDTO::getCode, m -> m, (k1, k2) -> k1));
        }
        // 从第5行开始写数据
        HSSFRow dataRow;
        for (int i = 4; i < lt1.size() + 4; i++) {
            dataRow = sheet.createRow(i);
            ChainCheckReportDTO chainCheckReportDTO =  lt1.get(i - 4);
            // 加入核酸检测数据
            if (null != map.get(chainCheckReportDTO.getCode())){
                ChainCheckReportDTO dto = map.get(chainCheckReportDTO.getCode());
                chainCheckReportDTO.setHsTodayCount(dto.getHsTodayCount());
                chainCheckReportDTO.setHsTotalCount(dto.getHsTotalCount());
                chainCheckReportDTO.setHsGT60(dto.getHsGT60());
            }else{
                chainCheckReportDTO.setHsTodayCount(0);
                chainCheckReportDTO.setHsTotalCount(0);
                chainCheckReportDTO.setHsGT60(0);
            }
            int j = 0;
            setStringValue(dataRow,j++,chainCheckReportDTO.getAreaName(),headStyle);
            setStringValue(dataRow,j++,chainCheckReportDTO.getCompanyName(),headStyle);
            setStringValue(dataRow,j++,chainCheckReportDTO.getRiskType(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getChainOnboardStaffCount(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getHsTodayCount(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getHsTotalCount(),headStyle);
            setStringValue(dataRow,j++,chainCheckReportDTO.getHsCheckRate(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getHsGT60(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getChainStaffCount(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getVaccTodayFirstCountT1(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getVaccTotalFirstCountT1(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getVaccTodayFirstCountT2(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getVaccTodaySecondCountT2(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getVaccTotalFirstCountT2(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getVaccTotalSecondCountT2(),headStyle);
            setStringValue(dataRow,j++,chainCheckReportDTO.getVaccSecondRateT2(),headStyle);
            setStringValue(dataRow,j++,chainCheckReportDTO.getVaccRate(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getVaccBoosterCount(),headStyle);
            setIntValue(dataRow,j++,chainCheckReportDTO.getVaccBoosterDone(),headStyle);
            setStringValue(dataRow,j++,chainCheckReportDTO.getVaccBoosterRate(),headStyle);
        }
        int next = 4 + lt1.size();
        HSSFRow row5 = sheet.createRow(next);
        HSSFCell cell50 = row5.createCell(0);
        cell50.setCellValue("合计：");
        cell50.setCellStyle(headStyle);
        ChainCheckReportDTO totalDto = new ChainCheckReportDTO();
        Integer chainStaffCount = lt1.stream().map(ChainCheckReportDTO::getChainStaffCount).reduce(0, Integer::sum);
        Integer chainOnboardStaffCount = lt1.stream().map(ChainCheckReportDTO::getChainOnboardStaffCount).reduce(0, Integer::sum);
        Integer hsTodayCount = lt1.stream().map(ChainCheckReportDTO::getHsTodayCount).reduce(0, Integer::sum);
        Integer hsTotalCount = lt1.stream().map(ChainCheckReportDTO::getHsTotalCount).reduce(0, Integer::sum);
        Integer HsGT60 = lt1.stream().map(ChainCheckReportDTO::getHsGT60).reduce(0, Integer::sum);
        Integer vaccTodayFirstCountT1 = lt1.stream().map(ChainCheckReportDTO::getVaccTodayFirstCountT1).reduce(0, Integer::sum);
        Integer vaccTotalFirstCountT1 = lt1.stream().map(ChainCheckReportDTO::getVaccTotalFirstCountT1).reduce(0, Integer::sum);

        Integer vaccTodayFirstCountT2 = lt1.stream().map(ChainCheckReportDTO::getVaccTodayFirstCountT2).reduce(0, Integer::sum);
        Integer vaccTodaySecondCountT2 = lt1.stream().map(ChainCheckReportDTO::getVaccTodaySecondCountT2).reduce(0, Integer::sum);
        Integer vaccTotalFirstCountT2 = lt1.stream().map(ChainCheckReportDTO::getVaccTotalFirstCountT2).reduce(0, Integer::sum);
        Integer vaccTotalSecondCountT2 = lt1.stream().map(ChainCheckReportDTO::getVaccTotalSecondCountT2).reduce(0, Integer::sum);
        Integer vaccBoosterCount = lt1.stream().map(ChainCheckReportDTO::getVaccBoosterCount).reduce(0, Integer::sum);
        Integer vaccBoosterDone = lt1.stream().map(ChainCheckReportDTO::getVaccBoosterDone).reduce(0, Integer::sum);
        totalDto.setHsGT60(HsGT60);
        totalDto.setHsTotalCount(hsTotalCount);
        totalDto.setHsTodayCount(hsTodayCount);
        totalDto.setChainStaffCount(chainStaffCount);
        totalDto.setChainOnboardStaffCount(chainOnboardStaffCount);
        totalDto.setVaccTodayFirstCountT1(vaccTodayFirstCountT1);
        totalDto.setVaccTotalFirstCountT1(vaccTotalFirstCountT1);

        totalDto.setVaccTodayFirstCountT2(vaccTodayFirstCountT2);
        totalDto.setVaccTodaySecondCountT2(vaccTodaySecondCountT2);
        totalDto.setVaccTotalFirstCountT2(vaccTotalFirstCountT2);
        totalDto.setVaccTotalSecondCountT2(vaccTotalSecondCountT2);
        totalDto.setVaccBoosterCount(vaccBoosterCount);
        totalDto.setVaccBoosterDone(vaccBoosterDone);

        int j = 3;
        setIntValue(row5,j++,totalDto.getChainOnboardStaffCount(),headStyle);
        setIntValue(row5,j++,totalDto.getHsTodayCount(),headStyle);
        setIntValue(row5,j++,totalDto.getHsTotalCount(),headStyle);
        setStringValue(row5,j++,totalDto.getHsCheckRate(),headStyle);
        setIntValue(row5,j++,totalDto.getHsGT60(),headStyle);
        setIntValue(row5,j++,totalDto.getChainStaffCount(),headStyle);
        setIntValue(row5,j++,totalDto.getVaccTodayFirstCountT1(),headStyle);
        setIntValue(row5,j++,totalDto.getVaccTotalFirstCountT1(),headStyle);
        setIntValue(row5,j++,totalDto.getVaccTodayFirstCountT2(),headStyle);
        setIntValue(row5,j++,totalDto.getVaccTodaySecondCountT2(),headStyle);
        setIntValue(row5,j++,totalDto.getVaccTotalFirstCountT2(),headStyle);
        setIntValue(row5,j++,totalDto.getVaccTotalSecondCountT2(),headStyle);
        setStringValue(row5,j++,totalDto.getVaccSecondRateT2(),headStyle);
        setStringValue(row5,j++,totalDto.getVaccRate(),headStyle);
        setIntValue(row5,j++,totalDto.getVaccBoosterCount(),headStyle);
        setIntValue(row5,j++,totalDto.getVaccBoosterDone(),headStyle);
        setStringValue(row5,j++,totalDto.getVaccBoosterRate(),headStyle);

        CellRangeAddress craSum = new CellRangeAddress(next, next, 0, 2);
        sheet.addMergedRegion(craSum);
        setBorder(craSum,sheet);
        File file = File.createTempFile("shcc-report-hsjc", ".xls");
        try (OutputStream fileOut = new FileOutputStream(file)) {
            workbook.write(fileOut);   //将workbook写入文件流
        } finally {
            workbook.close();
        }

        return new FileInputStream(file);
    }

    private void setBorder(CellRangeAddress region, Sheet sheet){
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet); // 下边框
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet); // 左边框
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet); // 有边框
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet); // 上边框
    }

    public void setStringValue( HSSFRow dataRow,int i,String value, CellStyle headStyle){
        HSSFCell cell = dataRow.createCell(i);
        cell.setCellValue(value);
        cell.setCellStyle(headStyle);
    }

    public void setIntValue( HSSFRow dataRow,int i,Integer value, CellStyle headStyle){
        HSSFCell cell = dataRow.createCell(i);
        cell.setCellValue(value);
        cell.setCellStyle(headStyle);
    }
}
