package com.mz.poi.mapper.annotation;

public enum Match {
  /**
   * 엑셀 > 모델 Row 변환시, 모든 필드의 값이 존재하면 파싱이 가능
   */
  ALL,
  /**
   * 엑셀 > 모델 Row 변환시, 모든 필수값 필드의 값이 존재하면 파싱이 가능
   */
  REQUIRED;
}
