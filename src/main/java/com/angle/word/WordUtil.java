package com.angle.word;

import java.util.List;

import javax.xml.bind.JAXBElement;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.dml.chart.CTBarChart;
import org.docx4j.dml.chart.CTBarSer;
import org.docx4j.dml.chart.CTNumVal;
import org.docx4j.wml.*;

import com.google.common.collect.Lists;

/**
 * Created by liupei on 2018/8/31.
 */
public class WordUtil {

    /**
     * 替换BarChart内容
     *
     * @param ctBarChart
     * @param values
     */
    public static void setBarChart(CTBarChart ctBarChart, List<String> values) {
        List<CTBarSer> ctBarSerList = ctBarChart.getSer();
        if (CollectionUtils.isNotEmpty(ctBarSerList)) {
            CTBarSer ctBarSer = ctBarSerList.get(0);
            List<CTNumVal> ctNumValList = ctBarSer.getVal().getNumRef().getNumCache().getPt();
            if (CollectionUtils.isNotEmpty(ctNumValList)) {
                for (int i = 0; i < ctNumValList.size(); i++) {
                    CTNumVal ctNumVal = ctNumValList.get(i);
                    ctNumVal.setV(values.get(i));
                }
            }
        }
    }

    /**
     * Description:设置单元格内容
     */
    public static void setTcContent(Tc tc, String content, RPr rpr) {
        List<Object> pList = tc.getContent();
        P p = null;
        if (CollectionUtils.isNotEmpty(pList)) {
            if (pList.get(0) instanceof P) {
                p = (P) pList.get(0);
            }
        } else {
            p = new P();
            tc.getContent().add(p);
        }
        R run = new R();
        p.getContent().add(run);
        run.setRPr(rpr);
        if (StringUtils.isNotBlank(content)) {
            Text text = new Text();

            // 清除获取单元格的内容
            p.getContent().clear();
            p.getContent().add(run);
            text.setSpace("preserve");
            // 设置单元格中的值
            text.setValue(content);
            run.getContent().add(text);
        }

    }

    /**
     * @Description: 在表格指定位置新增一行(默认按表格定义的列数添加)
     */
    public static Tr addTrByIndex(Tbl tbl, int index, TcPr tcPr) {
        TblGrid tblGrid = tbl.getTblGrid();
        Tr tr = new Tr();
        if (tblGrid != null) {
            List<TblGridCol> gridList = tblGrid.getGridCol();
            for (int i = 0; i < gridList.size(); i++) {
                Tc tc = new Tc();
                P p = new P();
                if (i == 0) {
                    // 设置新增行的第一列格式
                    tc.setTcPr(tcPr);

                }
                tc.getContent().add(p);
                tr.getContent().add(tc);

            }
        }
        if (index >= 0 && index < tbl.getContent().size()) {
            tbl.getContent().add(index, tr);
        } else {
            tbl.getContent().add(tr);
        }
        return tr;
    }

    /**
     * Description: 获取所有的单元格
     */
    public static List<Tc> getTrAllCell(Tr tr) {
        List<Object> objList = getAllElementFromObject(tr, Tc.class);
        List<Tc> tcList = Lists.newArrayList();
        if (objList == null) {
            return tcList;
        }
        for (Object tcObj : objList) {
            if (tcObj instanceof Tc) {
                Tc objTc = (Tc) tcObj;
                tcList.add(objTc);
            }
        }
        return tcList;
    }

    /**
     * @Description: 得到表格所有的行
     */
    public static List<Tr> getTblAllTr(Tbl tbl) {
        List<Object> objList = getAllElementFromObject(tbl, Tr.class);
        List<Tr> trList = Lists.newArrayList();
        if (objList == null) {
            return trList;
        }
        for (Object obj : objList) {
            if (obj instanceof Tr) {
                Tr tr = (Tr) obj;
                trList.add(tr);
            }
        }
        return trList;

    }

    /**
     * 得到指定类型的元素
     *
     * @param obj
     * @param toSearch
     * @return
     */
    public static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = Lists.newArrayList();
        if (obj instanceof JAXBElement) {
            obj = ((JAXBElement<?>) obj).getValue();
        }
        if (obj.getClass().equals(toSearch)) {
            result.add(obj);
        } else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }
}
