package org.gymbrain.gymbrainexcel.annotations;

import java.lang.annotation.*;

@Target({ElementType.TYPE, ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelField {
    String name() default "";

    int index() default 0;
}
