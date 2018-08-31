package com.angle.word;

import static com.angle.word.WordUtil.*;

import java.io.File;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.XmlUtils;
import org.docx4j.dml.chart.CTBarChart;
import org.docx4j.dml.chart.CTLineChart;
import org.docx4j.dml.chart.CTLineSer;
import org.docx4j.dml.chart.CTNumVal;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.EmbeddedPackagePart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.utils.BufferUtil;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.Row;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;

/**
 * 根据已有的word模板文件，替换文件内容来生成新的word文件。</br>
 * 本word文档包含：1. 表格 2. 图表
 */
public class WordCreate {

    private static final Logger logger      = LoggerFactory.getLogger(WordCreate.class);

    /**
     * word中的表格，表格是可以直接定位的，在word文档中，选中表格，"插入"-"书签"，写上书签名称"T1",这样就可以直接根据书签名称来定位这个表格
     */
    static final String         TABLE1      = "//w:bookmarkStart[@w:name='T1']/ancestor::w:tbl";

    /**
     * word中第一个图表
     */
    static final String         CHART_NAME1 = "/word/charts/chart1.xml";

    /**
     * word中第一个图表对应的excel
     */
    static final String         EXCEL_NAME1 = "/word/embeddings/Microsoft_Excel____1.xlsx";

    public static void main(String[] args) throws Exception {

        // step1 加载模板文档template.docx
        WordprocessingMLPackage template;
        try {
            InputStream inputStream = WordCreate.class.getResourceAsStream("/template.docx");
            template = WordprocessingMLPackage.load(inputStream);

            // step2 替换标题(其他word文档中的变量都可以这么替换)
            MainDocumentPart documentPart = template.getMainDocumentPart();
            VariablePrepare.prepare(template);
            Map<String, String> titleMap = Maps.newHashMap();
            titleMap.put("title", "测试");
            documentPart.variableReplace(titleMap);

            // step3 渲染表格内容
            setTable(template);

            // step4 渲染图表数据
            HashMap<PartName, Part> parts = template.getParts().getParts();
            EmbeddedPackagePart epp = (EmbeddedPackagePart) parts.get(new PartName(EXCEL_NAME1));
            Chart chart = (Chart) parts.get(new PartName(CHART_NAME1));
            // 图表数据
            List<String> counts = Lists.newArrayList();
            counts.add("389");
            counts.add("478");
            counts.add("231");
            counts.add("897");
            List<String> percents = Lists.newArrayList();
            percents.add("0.195");
            percents.add("0.2396");
            percents.add("0.1158");
            percents.add("0.4496");
            // 渲染图表excel
            setExcel(epp, counts, percents);
            // 渲染图表
            setChart(chart, counts, percents);

            // step5 保存word文档
            String outputFile = "output.docx";
            template.save(new File(outputFile));

        } catch (Docx4JException e) {
            logger.error("word create failed.", e);

        }

    }

    /**
     * 替换word图表的excel数据
     *
     * @param epp
     * @param counts
     * @param percents
     * @throws Docx4JException
     */
    private static void setExcel(EmbeddedPackagePart epp, List<String> counts, List<String> percents) {
        {
            InputStream is = BufferUtil.newInputStream(epp.getBuffer());
            SpreadsheetMLPackage spreadSheet;
            try {
                spreadSheet = SpreadsheetMLPackage.load(is);
            } catch (Docx4JException e) {
                logger.error("excel load failed.", e);
                return;
            }
            Map<PartName, Part> partsMap = spreadSheet.getParts().getParts();
            for (Map.Entry<PartName, Part> parts2 : partsMap.entrySet()) {
                if (partsMap.get(parts2.getKey()) instanceof WorksheetPart) {
                    WorksheetPart wsp = (WorksheetPart) partsMap.get(parts2.getKey());
                    List<Row> rows = wsp.getJaxbElement().getSheetData().getRow();
                    for (int i = 1; i < rows.size(); i++) {
                        Row row = rows.get(i);
                        List<Cell> cells = row.getC();
                        if (cells.size() != 3) {
                            break;
                        }
                        cells.get(1).setV(counts.get(i - 1));
                        cells.get(2).setV(percents.get(i - 1));
                    }
                }
            }

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            SaveToZipFile staf = new SaveToZipFile(spreadSheet);
            try {
                staf.save(baos);
            } catch (Docx4JException e) {
                logger.error("excel save failed.", e);
                return;
            }
            epp.setBinaryData(baos.toByteArray());
        }
    }

    /**
     * 替换word图表的chart数据
     *
     * @param chart
     * @param counts
     * @param percents
     */
    private static void setChart(Chart chart, List<String> counts, List<String> percents) {
        {

            List<Object> objects;
            try {
                objects = chart.getContents().getChart().getPlotArea().getAreaChartOrArea3DChartOrLineChart();

                if (objects.size() >= 1) {
                    for (Object obj : objects) {
                        if (obj instanceof CTBarChart) {
                            CTBarChart ctBarChart = (CTBarChart) obj;
                            setBarChart(ctBarChart, counts);
                        }
                        if (obj instanceof CTLineChart) {
                            CTLineChart ctLineChart = (CTLineChart) obj;
                            List<CTLineSer> ctLineSerList = ctLineChart.getSer();
                            if (CollectionUtils.isNotEmpty(ctLineSerList)) {
                                CTLineSer ctLineSer = ctLineSerList.get(0);
                                List<CTNumVal> ctNumValList = ctLineSer.getVal().getNumRef().getNumCache().getPt();
                                for (int i = 0; i < ctNumValList.size(); i++) {
                                    CTNumVal ctNumVal = ctNumValList.get(i);
                                    ctNumVal.setV(percents.get(i));
                                }
                            }

                        }
                    }
                }
            } catch (Docx4JException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 渲染表格内容，新增行，并设置单元格内容
     * 
     * @param template
     * @throws Exception
     */
    private static void setTable(WordprocessingMLPackage template) throws Exception {

        // 表格数据
        List<String> trDataList = Lists.newArrayList();
        List<String> trData = Lists.newArrayList();
        trData.add("张三");
        trData.add("男");
        trData.add("35");

        List<String> trData2 = Lists.newArrayList();
        trData2.add("李四");
        trData2.add("女");
        trData2.add("15");
        trDataList.add(StringUtils.join(trData, ","));
        trDataList.add(StringUtils.join(trData2, ","));

        // 填充表格
        List nodes = template.getMainDocumentPart().getJAXBNodesViaXPath(TABLE1, false);
        Tbl table = (Tbl) XmlUtils.unwrap(nodes.get(0));
        setTable(table, trDataList);

    }

    private static void setTable(Tbl tbl, List<String> trValues) {
        List<Tr> trList = getTblAllTr(tbl);

        if (CollectionUtils.isEmpty(trList) || CollectionUtils.isEmpty(trValues)) {
            return;
        }
        int trSize = trList.size();

        // 获取行首列的单元格格式
        Tr tr = trList.get(0);
        Tc tc = getTrAllCell(tr).get(0);
        TcPr tcPr = tc.getTcPr();

        Tc tc1 = getTrAllCell(tr).get(1);
        R run = null;
        List<Object> rList = getAllElementFromObject(tc1, R.class);
        if (rList != null) {
            for (Object obj : rList) {
                if (obj instanceof R) {
                    run = (R) obj;
                    break;
                }
            }
        }
        // 数据格式
        RPr rpr = run.getRPr();
        for (String trValue : trValues) {
            String[] trValueAry = trValue.split(",");
            // 新增行
            Tr tr2 = addTrByIndex(tbl, trSize, tcPr);
            List<Tc> tcList = getTrAllCell(tr2);
            for (int i = 0; i < tcList.size(); i++) {
                Tc tc2 = tcList.get(i);
                setTcContent(tc2, trValueAry[i], rpr);
            }
            trSize++;
        }

    }

}
