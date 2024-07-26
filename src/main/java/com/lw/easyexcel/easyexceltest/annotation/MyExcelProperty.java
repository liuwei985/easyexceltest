package com.lw.easyexcel.easyexceltest.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Description:
 * @ClassName: MyExcelProperty
 * @Author: lw
 * @Date: 2024/7/5 16:04
 * @Version: 1.0
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface MyExcelProperty {
    String name() default "";
    int index() default Integer.MAX_VALUE;
}
