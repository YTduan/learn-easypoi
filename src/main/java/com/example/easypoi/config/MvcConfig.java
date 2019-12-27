package com.example.easypoi.config;

import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.ResourceHandlerRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurerAdapter;

/**
 * @author duanyuantong
 * @version Id: MvcConfig, v 0.1 2019-11-26 11:00 duanyuantong Exp $
 */
@Configuration
public class MvcConfig extends WebMvcConfigurerAdapter {
    @Override
    public void addResourceHandlers(ResourceHandlerRegistry registry) {
        // 这里之所以多了一"/",是为了解决打war时访问不到问题
        registry.addResourceHandler("/**").addResourceLocations("/", "classpath:/");
    }
}
