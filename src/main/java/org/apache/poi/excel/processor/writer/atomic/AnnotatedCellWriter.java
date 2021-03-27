package org.apache.poi.excel.processor.writer.atomic;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.ZonedDateTime;
import java.util.Date;
import java.util.function.Function;

import org.apache.commons.lang.ObjectUtils;
import org.apache.poi.excel.processor.reader.FieldReader;
import org.apache.poi.excel.utility.DateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class AnnotatedCellWriter extends FieldReader {
	private static final  Logger log = LoggerFactory.getLogger(AnnotatedCellWriter.class);

	private Function<Object, Object> numericConverter;

	private Function<Object, Date> dateConverter;

	public AnnotatedCellWriter(Field field) {
		super(field);
	}

	public void initDateConverter() {

		dateConverter = (Object obj) ->{
			if (field.getType() == Date.class) {
				return this.getValue(obj,Date.class);
			} else if (field.getType() == LocalDate.class) {
				return DateUtil.asDate(this.getValue(obj,LocalDate.class));
			} else if (field.getType() == LocalDateTime.class) {
				return DateUtil.asDate(this.getValue(obj,LocalDateTime.class));
			} else if (field.getType() == OffsetDateTime.class) {
				return DateUtil.asDate(this.getValue(obj,OffsetDateTime.class));
			} else if (field.getType() == ZonedDateTime.class) {
				return DateUtil.asDate(this.getValue(obj,ZonedDateTime.class));
			} else {
				return DateUtil.parse(ObjectUtils.defaultIfNull(obj, "").toString());
			}
		};
	}


	public <T> void initNumericConverter(Class<T> clazz) {
		numericConverter = (Object obj) -> this.getValue(obj,clazz);

	}
	public void  writeNumeric(Cell cell, Object obj) {
		Object genericObject = numericConverter.apply(obj);
		if (genericObject == null) {
			log.debug("An Excel of numeric cell family is null. Not writing anything. Cell: {} " , cell);
		}
		try {
			cell.setCellValue(Double.parseDouble(String.valueOf(genericObject)));
		} catch (Exception cce) {
			log.debug("An Excel cell is not recognized as Integer. Not writing anything in this cell: {}" , cell);
		}
	}

	public void writeDate(Cell cell, Object obj) {
		try {
			Date value = dateConverter.apply(obj);
			cell.setCellValue(value);
		} catch (IllegalArgumentException e) {
			log.warn("Unable to write Date to an Excel cell : {}. Defaulting to blank.",cell);
		}
	}
}
