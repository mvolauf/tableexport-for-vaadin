package com.vaadin.addon.tableexport;

import java.io.Serializable;
import java.util.Collection;
import java.util.List;

/**
 * @author thomas
 */
public interface TableHolder<T> extends Serializable {

    List<String> getColumnIds();

    String getColumnHeader(String columnId);

    Short getCellAlignment(String columnId);

    boolean isColumnCollapsed(String columnId);
    
    Class<?> getColumnType(String columnId);

    Object getColumnValue(T item, String columnId);
    
    Collection<T> getItems();

    int size();

    boolean isHierarchical();

    void setHierarchical(final boolean hierarchical);

    Collection<T> getChildren(T rootItem);
    
    Collection<T> getRootItems();
    
}
