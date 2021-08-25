package com.easyexcel.annotation;

import java.lang.annotation.*;

/**
 * @author wrh
 * @date 2021/8/25
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Inherited
public @interface ExcelColumn {

    /**
     * excel 列名称
     *
     * @return 列名称
     */
    String value();

    /**
     * 列下标
     * 用于导出, 指定列的顺序, 也用于翻转后行的顺序
     *
     * @return 列下标
     */
    int columnIdx() default 0;

}
