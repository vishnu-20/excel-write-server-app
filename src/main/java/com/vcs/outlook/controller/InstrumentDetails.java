package com.vcs.outlook.controller;

import java.time.LocalDate;
import java.util.ArrayList;

public class InstrumentDetails {
    private int indexId;
    private String name;
    private String type;
    private double price;
    private LocalDate date;

    public InstrumentDetails(int indexId, String name, String type, double price, LocalDate date) {
        this.indexId = indexId;
        this.name = name;
        this.type = type;
        this.price = price;
        this.date = date;
    }

    public int getIndexId() {
        return indexId;
    }

    public String getName() {
        return name;
    }

    public String getType() {
        return type;
    }

    public double getPrice() {
        return price;
    }

    public LocalDate getDate() {
        return date;
    }
    

    
}