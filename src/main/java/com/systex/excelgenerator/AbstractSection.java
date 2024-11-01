package com.systex.excelgenerator;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.Collection;

public abstract class AbstractSection<T> implements Section<T> {

    protected final Collection<T> data;

    protected abstract Row getHeaderRow();
    protected abstract Row getDataRow();
    protected abstract Row getFooterRow();

    public AbstractSection(Collection<T> data) {
        this.data = data;
    }

    @Override
    public boolean populate(){
        Row headerRow = getHeaderRow();
        Row dataRow = getDataRow();
        Row footerRow = getFooterRow();

        return false;
    }

    protected abstract Cell generateCell();
}
