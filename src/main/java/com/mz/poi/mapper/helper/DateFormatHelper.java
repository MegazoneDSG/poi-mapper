package com.mz.poi.mapper.helper;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;
import org.springframework.util.StringUtils;

public class DateFormatHelper {

  public static Date getDate(LocalDate localDate, String zoneId) {
    ZoneId zone;
    if (StringUtils.isEmpty(zoneId)) {
      zone = ZoneId.systemDefault();
    } else {
      zone = ZoneId.of("Asia/Seoul");
    }
    return Date.from(localDate.atStartOfDay(zone).toInstant());
  }

  public static Date getDate(LocalDateTime localDateTime, String zoneId) {
    ZoneId zone;
    if (StringUtils.isEmpty(zoneId)) {
      zone = ZoneId.systemDefault();
    } else {
      zone = ZoneId.of("Asia/Seoul");
    }
    return Date.from(localDateTime.atZone(zone).toInstant());
  }

  public static LocalDateTime getLocalDateTime(Date date, String zoneId) {
    ZoneId zone;
    if (StringUtils.isEmpty(zoneId)) {
      zone = ZoneId.systemDefault();
    } else {
      zone = ZoneId.of("Asia/Seoul");
    }
    return LocalDateTime.ofInstant(date.toInstant(), zone);
  }

  public static LocalDate getLocalDate(Date date, String zoneId) {
    return getLocalDateTime(date, zoneId).toLocalDate();
  }
}
