package com.example.easypoi.service.impl;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import com.example.easypoi.service.TestService;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author duanyuantong
 * @version Id: TestService, v 0.1 2019-11-26 10:53 duanyuantong Exp $
 */
@Service
public class TestServiceImpl implements TestService {


    @Override
    public  void test()  {


        File file1 = new File("/template.xlsx");
        String str = file1.getPath();
        TemplateExportParams params = new TemplateExportParams(str);
        System.out.println("sheet-----"+params.getSheetName());
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("purchaseNumber", "2332156464546");
        map.put("courier", "顺丰物流");
        map.put("logisticsNumber", "32165564456464");
        map.put("shouldNumber", "10");

        List<Map<String, String>> listMap = new ArrayList<Map<String, String>>();
        for (int i = 0; i < 4; i++) {
            Map<String, String> lm = new HashMap<String, String>();
            lm.put("id", i + 1 + "");
            lm.put("zijin", i * 10000 + "");
            lm.put("bianma", "A001");
            lm.put("mingcheng", "设计");
            lm.put("xiangmumingcheng", "EasyPoi " + i + "期");
            lm.put("quancheng", "开源项目");
            lm.put("sqje", i * 10000 + "");
            lm.put("hdje", i * 10000 + "");

            listMap.add(lm);
        }
        map.put("maplist", listMap);

        Workbook workbook = ExcelExportUtil.exportExcel(params, map);

        FileOutputStream fos = null;
        File file = new File("ceshi.xls");
        try {
            fos = new FileOutputStream(file,true);

            workbook.write(fos);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
