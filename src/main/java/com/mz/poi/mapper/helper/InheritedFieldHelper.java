package com.mz.poi.mapper.helper;

import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;

public class InheritedFieldHelper {

	public static Field[] getDeclaredFields(Class<?> clazz) {
		if (clazz == null) {
			return new Field[]{};
		}
		List<Field> declaredFields = new ArrayList<>(
				Arrays.asList(clazz.getDeclaredFields()));
		List<Field> superClassFields = Arrays.stream(getDeclaredFields(clazz.getSuperclass()))
				.filter(superClassField ->
						declaredFields.stream()
								.noneMatch(declaredField ->
										declaredField.getName().equals(superClassField.getName())))
				.collect(Collectors.toList());
		declaredFields.addAll(superClassFields);
		return declaredFields.toArray(new Field[0]);
	}

	public static Field getDeclaredField(Class<?> clazz, String fieldName)
			throws NoSuchFieldException {
		return Arrays.stream(getDeclaredFields(clazz))
				.filter(field -> field.getName().equals(fieldName))
				.findFirst()
				.orElseThrow(() -> new NoSuchFieldException(
						String.format("%s field not founded in %s class", fieldName,
								clazz.getName())));
	}

	public static Class<?> getGenericClass(Field field) {
		ParameterizedType genericType =
				(ParameterizedType) field.getGenericType();
		return (Class<?>) genericType.getActualTypeArguments()[0];
	}
}
