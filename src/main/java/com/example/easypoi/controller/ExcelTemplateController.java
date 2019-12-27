package com.example.easypoi.controller;

import com.example.easypoi.service.TestService;
import com.example.easypoi.service.impl.TestServiceImpl;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 * @author duanyuantong
 * @version Id: ExcelTemplateController, v 0.1 2019-11-25 22:52 duanyuantong Exp $
 */
@Controller
public class ExcelTemplateController {


    @Autowired
    private TestService testService;

    @RequestMapping("/excel")
    public  void test()  {
        testService.test();
    }


}
