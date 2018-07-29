package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Datavalidator {
  public static void validator(XSSFSheet worksheet, String value1, String value2, int firstcol) {
    XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(worksheet);
    XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
                                                        dvHelper.createExplicitListConstraint(new String[]{value1, value2});
    CellRangeAddressList addressList = new CellRangeAddressList(1, 3103, firstcol, firstcol);
    XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
            dvConstraint, addressList);
    validation.setShowErrorBox(true);
    worksheet.addValidationData(validation);
  }

  public static void validatorString(XSSFSheet worksheet,String[] args ,  int firstcol) {
    XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(worksheet);
    XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper.createExplicitListConstraint(args);
    CellRangeAddressList addressList = new CellRangeAddressList(1, 3103, firstcol, firstcol);
    XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, addressList);
    validation.setShowErrorBox(true);
    worksheet.addValidationData(validation);
  }

  public static void validatorType(XSSFSheet worksheet, String value1, String value2,String value3, int firstcol) {
    XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(worksheet);
    XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
                                                        dvHelper.createExplicitListConstraint(new String[]{value1,
                                                                value2,value3});
    CellRangeAddressList addressList = new CellRangeAddressList(1, 3103, firstcol, firstcol);
    XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
            dvConstraint, addressList);
    validation.setShowErrorBox(true);
    worksheet.addValidationData(validation);
  }

  public static void validatorState(XSSFSheet worksheet, String value1, String value2,String value3,String value4,int firstcol) {
    XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(worksheet);
    XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
                                                        dvHelper.createExplicitListConstraint(new String[]{value1,
                                                                value2,value3,value4});
    CellRangeAddressList addressList = new CellRangeAddressList(1, 3103, firstcol, firstcol);
    XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
            dvConstraint, addressList);
    validation.setShowErrorBox(true);
    worksheet.addValidationData(validation);
  }

  public static void validatorLedger(XSSFSheet worksheet, String value1, String value2,String value3,String value4,
                                     String value5,int firstcol) {
    XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(worksheet);
    XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
                                                        dvHelper.createExplicitListConstraint(new String[]{value1,
                                                                value2,value3,value4,value5});
    CellRangeAddressList addressList = new CellRangeAddressList(1, 3103, firstcol, firstcol);
    XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
            dvConstraint, addressList);
    validation.setShowErrorBox(true);
    worksheet.addValidationData(validation);
  }


  public static void validatorWeekday(XSSFSheet worksheet, String value1, String value2,String value3,String value4,
                                      String value5,String value6,String value7,int firstcol) {
    XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(worksheet);
    XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
                                                        dvHelper.createExplicitListConstraint(new String[]{value1,
                                                                value2,value3,value4,value5,value6,value7});
    CellRangeAddressList addressList = new CellRangeAddressList(1, 3103, firstcol, firstcol);
    XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
            dvConstraint, addressList);
    validation.setShowErrorBox(true);
    worksheet.addValidationData(validation);
  }
}
