package org.apache.poi.excel.processor.writer.atomic;

import java.lang.reflect.Field;
import java.util.Calendar;
import java.util.Date;
import java.util.function.BiConsumer;

import org.apache.poi.excel.model.ExcelCellType;
import org.apache.poi.excel.model.WorkbookContainer;
import org.apache.poi.excel.processor.reader.FieldReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * we will replace this multiple methods with the visitor design pattern so that we can reduce the complexity of the code
 */
public class GenericCellWriter extends FieldReader {
    private static final Logger log = LoggerFactory.getLogger(GenericCellWriter.class);
    private WorkbookContainer container;

    public GenericCellWriter(Field field, WorkbookContainer container) {
        super(field);
        this.container = container;
    }

	public BiConsumer<Cell, Object> writer = (Cell cell, Object obj) -> {
		if (obj instanceof Integer) {
			cell.setCellValue(this.getValue(obj, Integer.class));
			cell.setCellStyle(container.getStyle(ExcelCellType.INTEGER));
		} else if (obj instanceof Short) {
			cell.setCellValue(this.getValue(obj, Short.class));
                cell.setCellStyle(container.getStyle(ExcelCellType.INTEGER));
		} else if (obj instanceof Long) {
			cell.setCellValue(this.getValue(obj, Long.class));
			cell.setCellStyle(container.getStyle(ExcelCellType.INTEGER));

		} else if (obj instanceof Double) {
			cell.setCellValue(this.getValue(obj, Double.class));
			cell.setCellStyle(container.getStyle(ExcelCellType.PRECISE));

		} else if (obj instanceof Float) {
			cell.setCellValue(this.getValue(obj, Float.class));
			cell.setCellStyle(container.getStyle(ExcelCellType.DECIMAL));

		} else if (obj instanceof Byte) {
			cell.setCellValue(this.getValue(obj, Byte.class));
			cell.setCellStyle(container.getStyle(ExcelCellType.INTEGER));

		}
		else {
			throw new IllegalArgumentException("the argurment that you try to write has an unknown type");
		}
	};



    @SuppressWarnings("deprecation")
    public BiConsumer<Cell, Object> booleanWriter = (Cell cell, Object obj) -> {
        cell.setCellValue(this.getValue(obj, Boolean.class));
        cell.setCellType(CellType.BOOLEAN);
    };

    @SuppressWarnings("deprecation")
    public BiConsumer<Cell, Object> utilDateWriter = (Cell cell, Object obj) -> {
        Date value = null;
        try {
            value = (Date) field.get(obj);

        } catch (IllegalArgumentException | IllegalAccessException | NullPointerException | ClassCastException e) {
            log.warn("Unable to write cell : {} . Defaulting to ERROR.",cell);
            cell.setCellType(CellType.ERROR);
        }
        cell.setCellValue(value);
        cell.setCellStyle(container.getStyle(ExcelCellType.DATE));
    };

    @SuppressWarnings("deprecation")
    public BiConsumer<Cell, Object> sqlDateWriter = (Cell cell, Object obj) -> {
        Date value = null;
        try {
            value = new Date(((java.sql.Date) field.get(obj)).getTime());
            cell.setCellStyle(container.getStyle(ExcelCellType.DATE));
        } catch (IllegalArgumentException | IllegalAccessException | NullPointerException | ClassCastException e) {
            log.warn("Unable to write cell : {} .  Defaulting to ERROR.",cell);
            cell.setCellType(CellType.ERROR);
        }
        cell.setCellValue(value);
    };

    @SuppressWarnings("deprecation")
    public BiConsumer<Cell, Object> calendarWriter = (Cell cell, Object obj) -> {
        Date value = null;
        try {
            value = ((Calendar) field.get(obj)).getTime();
            cell.setCellStyle(container.getStyle(ExcelCellType.DATETIME));
        } catch (IllegalArgumentException | IllegalAccessException | NullPointerException | ClassCastException e) {
            log.warn("Unable to write cell : {}. Defaulting to ERROR.",cell);
            cell.setCellType(CellType.ERROR);
        }
        cell.setCellValue(value);
    };

    @SuppressWarnings("deprecation")
    public BiConsumer<Cell, Object> stringWriter = (Cell cell, Object obj) -> {
        obj = this.getValue(obj, String.class);
        if (obj != null) {
            cell.setCellValue(obj.toString());
            cell.setCellType(CellType.STRING);
        }
    };
}
