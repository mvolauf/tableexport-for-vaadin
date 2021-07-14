package com.vaadin.addon.tableexport;

import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

import com.vaadin.data.HasHierarchicalDataProvider;
import com.vaadin.data.provider.HierarchicalQuery;
import com.vaadin.data.provider.Query;
import com.vaadin.ui.Grid;
import com.vaadin.ui.Grid.Column;
import com.vaadin.ui.renderers.Renderer;

public class DefaultGridHolder<T> implements TableHolder<T> {

	protected short defaultAlignment = HorizontalAlignment.LEFT.getCode();

	private boolean hierarchical = false;

	protected Grid<T> heldGrid;
	private List<String> columnIds;

	public DefaultGridHolder(Grid<T> grid) {
		this.heldGrid = grid;
		this.columnIds = heldGrid.getColumns().stream().map(Column::getId).collect(Collectors.toList());
		setHierarchical(grid instanceof HasHierarchicalDataProvider);
	}

	@Override
	public List<String> getColumnIds() {
		return columnIds;
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
	public Short getCellAlignment(String columnId) {
		if (null == heldGrid) {
			return defaultAlignment;
		}
		Renderer<?> renderer = getRenderer(columnId);
		if (renderer != null) {
			if (ExcelExport.isNumeric(renderer.getPresentationType())) {
				return HorizontalAlignment.RIGHT.getCode();
			}
		}
		return defaultAlignment;
	}

	@Override
	public boolean isColumnCollapsed(String columnId) {
		if (null == heldGrid) {
			return false;
		}
		return heldGrid.getColumn(columnId).isHidden();
	}

	@Override
	public String getColumnHeader(String columnId) {
		if (null != heldGrid) {
			Column<?, ?> c = getColumn(columnId);
			return c.getCaption();
		} else {
			return columnId.toString();
		}
	}

	protected Column<T, ?> getColumn(String columnId) {
		return heldGrid.getColumn(columnId);
	}

	protected Renderer<?> getRenderer(String columnId) {
		Column<T, ?> column = getColumn(columnId);
		if (column != null) {
			return column.getRenderer();
		}
		return null;
	}

	@Override
	public Class<?> getColumnType(String columnId) {
		Renderer<?> renderer = getRenderer(columnId);
		if (renderer != null) {
			return renderer.getPresentationType();
		} else {
			return String.class;
		}
	}

	@Override
	public Object getColumnValue(T item, String columnId) {
		return getColumn(columnId).getValueProvider().apply(item);
	}

	@Override
	public Collection<T> getChildren(T rootItem) {
		if (isHierarchical()) {
			return ((HasHierarchicalDataProvider<T>) heldGrid).getTreeData().getChildren(rootItem);
		} else {
			return Collections.emptyList();
		}
	}

	@Override
	public Collection<T> getItems() {
		return heldGrid.getDataProvider()
				.fetch(new Query<>(0, Integer.MAX_VALUE, heldGrid.getDataCommunicator().getBackEndSorting(),
						heldGrid.getDataCommunicator().getInMemorySorting(), null))
				.collect(Collectors.toList());
	}

    @Override
    public int size() {
      return heldGrid.getDataProvider().size(isHierarchical() ? new HierarchicalQuery<>(null, null) : new Query<>());
    }

	@Override
	public Collection<T> getRootItems() {
		if (isHierarchical()) {
			return ((HasHierarchicalDataProvider<T>) heldGrid).getTreeData().getRootItems();
		} else {
			return getItems();
		}
	}

}
