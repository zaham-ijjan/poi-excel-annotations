package org.apache.poi.excel.processor.reader;

import lombok.SneakyThrows;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;

/**
 * A generalization of the Reflection - translation utilities.
 * 
 * @author ssp5zone
 */
public class FieldReader {
	private static final  Logger log = LoggerFactory.getLogger(FieldReader.class);

	protected Field field;

	public FieldReader(Field field) {
		this.field = field;
	}



	/**
	 * this method is more generic than the previous ones it's much easier to specify the class type
	 * and cast it instead of creating for each type his own class as for the exception ,it will handled automatically with
	 * @Sneakythrows since it's a runtime Exception
	 * @param obj
	 * @param clazz
	 * @param <T>
	 * @return
	 */
	@SneakyThrows
	public <T> T getValue(Object obj , Class<T> clazz){
		return clazz.cast(field.get(obj));
	}

}
