package com.mz.poi.mapper.helper;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class InheritedFieldHelper {

  public static Field[] getDeclaredFields(Class<?> clazz) {
    if (clazz == null) {
      return new Field[]{};
    }
    List<Field> superClassFields = new ArrayList<>(
        Arrays.asList(getDeclaredFields(clazz.getSuperclass())));
    List<Field> declaredFields = new ArrayList<>(
        Arrays.asList(clazz.getDeclaredFields()));
    declaredFields.addAll(superClassFields);
    return declaredFields.toArray(new Field[0]);
  }

  public static Field getDeclaredField(Class<?> clazz, String fieldName)
      throws NoSuchFieldException {
    return Arrays.stream(getDeclaredFields(clazz))
        .filter(field -> field.getName().equals(fieldName))
        .findFirst()
        .orElseThrow(() -> new NoSuchFieldException(
            String.format("%s field not founded in %s class", fieldName, clazz.getName())));
  }
}
