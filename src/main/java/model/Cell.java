package model;

import java.util.List;

public class Cell {

    private float value;
    private List<Cell> in;

    public float getValue() {
        return value;
    }

    public void setValue(float value) {
        this.value = value;
    }

    public List<Cell> getIn() {
        return in;
    }

    public void setIn(List<Cell> in) {
        this.in = in;
    }
}
