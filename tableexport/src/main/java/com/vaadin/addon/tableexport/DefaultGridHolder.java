package com.vaadin.addon.tableexport;

import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

import com.vaadin.data.HasHierarchicalDataProvider;
import com.vaadin.data.ValueProvider;
import com.vaadin.data.provider.Query;
import com.vaadin.ui.Grid;
import com.vaadin.ui.Grid.Column;
import com.vaadin.ui.renderers.Renderer;

public class DefaultGridHolder<T> implements TableHolder {

    protected short defaultAlignment = HorizontalAlignment.LEFT.getCode();

    private boolean hierarchical = false;

    protected Grid<T> heldGrid;
    private List<String> propIds;

    public DefaultGridHolder(Grid<T> grid) {
        this.heldGrid = grid;
        this.propIds = heldGrid.getColumns().stream().map(Column::getId).collect(Collectors.toList());
        setHierarchical(grid instanceof HasHierarchicalDataProvider);
    }

    @Override
    public List<String> getPropIds() {
        return propIds;
    }

    @Override
    public boolean isHierarchical() {
        return hierarchical;
    }

    @Override
    final public void setHierarchical(boolean hierarchical) {
        this.hierarchical = hierarchical;
    }

    @Override
    public Short getCellAlignment(Object propId) {
        if (null == heldGrid) {
            return defaultAlignment;
        }
        Renderer<?> renderer = getRenderer(propId);
        if (renderer != null) {
            if (ExcelExport.isNumeric(renderer.getPresentationType())) {
            	return HorizontalAlignment.RIGHT.getCode();
            }
        }
        return defaultAlignment;
    }

    @Override
    public boolean isGeneratedColumn(final Object propId) throws IllegalArgumentException {
        return false;
    }

    @Override
    public Class<?> getPropertyTypeForGeneratedColumn(final Object propId) throws IllegalArgumentException {
        throw new UnsupportedOperationException();
    }

    @Override
    public boolean isExportableFormattedProperty() {
        return false;
    }

    @Override
    public boolean isColumnCollapsed(Object propertyId) {
        if (null == heldGrid) {
            return false;
        }
        return heldGrid.getColumn((String) propertyId).isHidden();
    }

    @Override
    public String getColumnHeader(Object propertyId) {
        if (null != heldGrid) {
            Column<?,?> c = getColumn(propertyId);
            return c.getCaption();
        } else {
            return propertyId.toString();
        }
    }

    protected Column<T, ?> getColumn(Object propId) {
    	return heldGrid.getColumn((String) propId);
    }

    protected Renderer<?> getRenderer(Object propId) {
    	Column<T, ?> column = getColumn(propId);
    	if (column != null) {
    		return column.getRenderer();
    	}
    	return null;
    }

    @Override
    public Class<?> getPropertyType(Object propId) {
        Renderer<?> renderer = getRenderer(propId);
        if (renderer != null) {
            return renderer.getPresentationType();
        } else {
            return String.class;
        }
    }

    @Override
    public Object getPropertyValue(Object itemId, Object propId, boolean useTableFormatPropertyValue) {
    	return getColumn((String)propId).getValueProvider().apply((T) itemId);
    }

    @Override
    public Collection<T> getChildren(Object rootItemId) {
    	if (isHierarchical()) {
    		return ((HasHierarchicalDataProvider<T>) heldGrid).getTreeData().getChildren((T)rootItemId);
        } else {
        	return Collections.emptyList();
        }
    }
    
    @Override
    public Collection<T> getItemIds() {
    	return heldGrid.getDataProvider().fetch(new Query<>(0, Integer.MAX_VALUE,
    	          heldGrid.getDataCommunicator().getBackEndSorting(), heldGrid.getDataCommunicator().getInMemorySorting(), null)).collect(Collectors.toList());
    }

    @Override
    public Collection<T> getRootItemIds() {
    	if (isHierarchical()) {
    		return ((HasHierarchicalDataProvider<T>) heldGrid).getTreeData().getRootItems();
    	} else {
    		return getItemIds();
    	}
    }

}
