package model;

import java.util.List;

public class Spreadsheet {

    private String name;
    private List<Cell> hasCells;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Cell> getHasCells() {
        return hasCells;
    }

    public void setHasCells(List<Cell> hasCells) {
        this.hasCells = hasCells;
    }
}
