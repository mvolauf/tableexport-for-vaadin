package com.vaadin.addon.tableexport;

import java.io.Serializable;
import java.util.Collection;
import java.util.List;

/**
 * @author thomas
 */
public interface TableHolder extends Serializable {

    List<Object> getPropIds();
    boolean isHierarchical();
    void setHierarchical(final boolean hierarchical);

    Short getCellAlignment(Object propId);
    boolean isGeneratedColumn(final Object propId) throws IllegalArgumentException;
    Class<?> getPropertyTypeForGeneratedColumn(final Object propId) throws IllegalArgumentException;

    // table delegated methods
    boolean isColumnCollapsed(Object propertyId);
    String getColumnHeader(Object propertyId);
    boolean isExportableFormattedProperty();
    
    Class<?> getPropertyType(Object propId);

    Object getPropertyValue(Object itemId, Object propId, boolean useTableFormatPropertyValue);

    Collection<?> getChildren(Object rootItemId);
    
    Collection<?> getItemIds();

    Collection<?> getRootItemIds();

}
