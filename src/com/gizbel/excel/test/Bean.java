package com.gizbel.excel.test;

import com.gizbel.excel.annotations.ExcelBean;
import com.gizbel.excel.annotations.ExcelColumnIndex;

@ExcelBean
public class Bean {

    @ExcelColumnIndex(columnIndex = "0", dataType = "double", defaultValue = "2.356")
    private Double fee;

    @ExcelColumnIndex(columnIndex = "1", dataType = "double")
    private Double totalCost;

    @ExcelColumnIndex(columnIndex = "2", dataType = "string")
    private String reference;

    @ExcelColumnIndex(columnIndex = "3", dataType = "string")
    private String invoiceNumber;

    public Double getFee() {
        return fee;
    }

    public void setFee(Double fee) {
        this.fee = fee;
    }

    public Double getTotalCost() {
        return totalCost;
    }

    public void setTotalCost(Double totalCost) {
        this.totalCost = totalCost;
    }

    public String getReference() {
        return reference;
    }

    public void setReference(String reference) {
        this.reference = reference;
    }

    public String getInvoiceNumber() {
        return invoiceNumber;
    }

    public void setInvoiceNumber(String invoiceNumber) {
        this.invoiceNumber = invoiceNumber;
    }

    @Override
    public String toString() {
        return ("Fees : " + this.fee + "\nTotal Cost : " + this.totalCost + "\nReference : " + this.reference
                + "\nInvoice Number : " + this.invoiceNumber);
    }

}
