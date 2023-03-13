package com.angle.word;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.io.output.ByteArrayOutputStream;
import org.docx4j.dml.chart.*;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io3.Save;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.SharedStrings;
import org.docx4j.openpackaging.parts.SpreadsheetML.TablePart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.EmbeddedPackagePart;
import org.docx4j.utils.BufferUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xlsx4j.sml.CTRst;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.Row;
import org.xlsx4j.sml.STCellType;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 根据已有的word模板文件，替换文件内容来生成新的word文件。</br>
 */
public class WordCreate {

    private static final Logger logger = LoggerFactory.getLogger(WordCreate.class);

    /**
     * 用于匹配excel中的范围字符串，例如：Sheet1!$A$1:$C$3
     */
    private static final Pattern PATTERN = Pattern.compile("[^!]+!\\$([A-Z]+)\\$([0-9]+):\\$([A-Z]+)\\$([0-9]+)");

    public static void main(String[] args) throws Exception {
        List<Map<String, Object>> data = new ObjectMapper().readValue(WordCreate.class.getResourceAsStream("/data.json"), new TypeReference<List<Map<String, Object>>>() {});
        InputStream inputStream = WordCreate.class.getResourceAsStream("/template.docx");
        try {
            generate(data, inputStream);
        } catch (Docx4JException e) {
            logger.error("word create failed.", e);
        }

    }

    private static void generate(List<Map<String, Object>> data, InputStream inputStream) throws Docx4JException, JAXBException {
        WordprocessingMLPackage template = WordprocessingMLPackage.load(inputStream);

        // 查找word中所有图表
        List<Object> chartElements = template.getMainDocumentPart().getJAXBNodesViaXPath("//c:chart", false);
        for (Object element : chartElements) {
            // 根据id获取图表对象
            String chartId = ((CTRelId) ((JAXBElement<?>) element).getValue()).getId();
            Chart chart = (Chart) template.getMainDocumentPart().getRelationshipsPart().getPart(chartId);

            // 获取内嵌的excel对象
            CTExternalData externalData = chart.getContents().getExternalData();
            EmbeddedPackagePart epp = (EmbeddedPackagePart) chart.getRelationshipsPart().getPart(externalData.getId());

            // 更新excel中的数据
            List<String> columns = updateExcel(epp, data);

            // 更新图表
            setChart(chart, data, columns);
        }

        // 保存word文档
        String outputFile = "output.docx";
        template.save(new File(outputFile));
    }

    /**
     * 更新excel中的数据，第一行为表头，有几列数据需要在模板中设置
     */
    private static List<String> updateExcel(EmbeddedPackagePart epp, List<Map<String, Object>> data) throws Docx4JException {
        // 加载excel
        InputStream is = BufferUtil.newInputStream(epp.getBuffer());
        SpreadsheetMLPackage spreadSheet = SpreadsheetMLPackage.load(is);

        List<String> columns = null;

        Map<PartName, Part> partsMap = spreadSheet.getParts().getParts();
        for (Part part : partsMap.values()) {
            if (part instanceof WorksheetPart) {
                WorksheetPart wsp = (WorksheetPart) part;
                List<Row> rows = wsp.getContents().getSheetData().getRow();

                // 获取表头
                Row headerRow = rows.get(0);
                columns = headerRow.getC().stream().map(cell -> getCellValue(cell, spreadSheet)).collect(Collectors.toList());

                // 获取内容的第一行
                Row contentRow = rows.get(1);

                rows.clear();
                rows.add(headerRow);

                for (int rowIndex = 0; rowIndex < data.size(); rowIndex += 1) {
                    Map<String, Object> item = data.get(rowIndex);
                    Row row = new Row();
                    List<Cell> cells = row.getC();

                    for (int columnIndex = 0; columnIndex < columns.size(); columnIndex += 1) {
                        Cell cell = new Cell();
                        // 数据从A列开始
                        cell.setR(String.valueOf((char) ('A' + columnIndex)) + (rowIndex + 2));

                        // 设置模板中该单元格的属性
                        Cell contentCell = contentRow.getC().get(columnIndex);
                        cell.setS(contentCell.getS());
                        cell.setT(contentCell.getT());

                        // 设置单元格的值
                        String column = columns.get(columnIndex);
                        Object value = item.get(column);
                        setCellValue(cell, value, spreadSheet);

                        cells.add(cell);
                    }
                    rows.add(row);
                }
            }
        }

        if (columns != null) {
            for (Part part : partsMap.values()) {
                if (part instanceof TablePart) {
                    // 更新表格区域
                    TablePart tablePart = (TablePart) part;
                    tablePart.getContents().setRef("A1:" + (char)((columns.size() - 1) + 'A') + (data.size() + 1));
                }
            }
        }

        // 保存Excel
        try {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            Save s = new Save(spreadSheet);
            s.save(bos);
            epp.setBinaryData(bos.toByteArray());
        } catch (Docx4JException e) {
            logger.error("excel save failed.", e);
            return null;
        }

        return columns;
    }

    /**
     * 获取单元格的值
     */
    private static String getCellValue(Cell cell, SpreadsheetMLPackage spreadSheet) {
        if (cell.getT() == STCellType.S) {
            // 查找共享字符串
            try {
                SharedStrings sharedStrings = (SharedStrings) spreadSheet.getParts().get(new PartName("/xl/sharedStrings.xml"));
                CTRst si = sharedStrings.getContents().getSi().get(Integer.parseInt(cell.getV()));
                return si.getT().getValue();
            } catch (Docx4JException e) {
                logger.error("get cell value", e);
            }
        }
        return cell.getV();
    }

    /**
     * 设置单元格的值
     */
    private static void setCellValue(Cell cell, Object value, SpreadsheetMLPackage spreadSheet) {
        if (value instanceof String) {
            try {
                // 查找共享字符串
                SharedStrings sharedStrings = (SharedStrings) spreadSheet.getParts().get(new PartName("/xl/sharedStrings.xml"));
                int index = getSharedStringIndex(sharedStrings.getContents().getSi(), value);
                if (index >= 0) {
                    cell.setT(STCellType.S);
                    cell.setV(String.valueOf(index));
                } else {
                    cell.setT(STCellType.STR);
                    cell.setV(value.toString());
                }
            } catch (Docx4JException e) {
                logger.error("get cell value", e);
            }
        } else {
            cell.setV(value == null ? null : value.toString());
        }
    }

    /**
     * 查找共享字符串下标
     */
    private static int getSharedStringIndex(List<CTRst> si, Object value) {
        for (int i = 0; i < si.size(); i++) {
            if (value.equals(si.get(i).getT().getValue())) {
                return i;
            }
        }
        return -1;
    }

    /**
     * 替换word图表的chart数据
     */
    private static void setChart(Chart chart, List<Map<String, Object>> data, List<String> columns) throws JAXBException, XPathBinderAssociationIsPartialException {
        // 替换数字引用
        List<Object> numRefList = chart.getJAXBNodesViaXPath("//c:numRef", false);
        for (Object element : numRefList) {
            CTNumRef ref = (CTNumRef) element;
            int[] address = getCellAddress(ref.getF());
            if (address == null) {
                continue;
            }

            List<CTNumVal> ptList = ref.getNumCache().getPt();
            ptList.clear();

            int index = 0;
            for (Map<String, Object> item : data) {
                Object value = item.get(columns.get(address[1]));
                CTNumVal val = new CTNumVal();
                val.setIdx(index++);
                val.setV(value == null ? null : value.toString());
                ptList.add(val);
            }

            // 根据数据长度重新计算f
            ref.setF(getNewF(ref.getF(), data.size()));
        }

        // 替换字符串引用
        List<Object> strRefList = chart.getJAXBNodesViaXPath("//c:strRef", false);
        for (Object element : strRefList) {
            CTStrRef ref = (CTStrRef) element;
            int[] address = getCellAddress(ref.getF());
            // 找不到单元格地址，或单元格为表头
            if (address == null || address[0] == 0) {
                continue;
            }

            List<CTStrVal> ptList = ref.getStrCache().getPt();
            ptList.clear();

            int index = 0;
            for (Map<String, Object> item : data) {
                Object value = item.get(columns.get(address[1]));
                CTStrVal val = new CTStrVal();
                val.setIdx(index++);
                val.setV(value == null ? null : value.toString());
                ptList.add(val);
            }

            // 根据数据长度重新计算f
            ref.setF(getNewF(ref.getF(), data.size()));
        }
    }

    /**
     * 把单元格字符串地址解析为数字地址
     */
    private static int[] getCellAddress(String f) {
        Matcher m = PATTERN.matcher(f);
        if (m.find()) {
            return new int[] { Integer.parseInt(m.group(2)) - 1, m.group(1).charAt(0) - 'A' };
        }
        return null;
    }

    /**
     * 根据数据长度重新计算f
     */
    private static String getNewF(String f, int size) {
        Matcher m = PATTERN.matcher(f);
        if (m.find()) {
            return f.substring(0,m.start(1)) + m.group(1)  + "$2:$" + m.group(3) + "$" + (size + 1);
        }
        return f;
    }
}
