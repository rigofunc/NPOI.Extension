// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    /// <summary>
    /// Typed row data validator delegate, validate current row before adding it to the result list.
    /// </summary>
    /// <param name="rowIndex">Index of current row in excel</param>
    /// <param name="rowData">Model data of current row</param>
    /// <returns>Whether the row data passes validation</returns>
    public delegate bool RowDataValidator<TModel>(int rowIndex, TModel rowData) where TModel : class;

    /// <summary>
    /// Row data validator delegate, validate current row before adding it to the result list.
    /// </summary>
    /// <param name="rowIndex">Index of current row in excel</param>
    /// <param name="rowData">Model data of current row</param>
    /// <returns>Whether the row data passes validation</returns>
    public delegate bool RowDataValidator(int rowIndex, object rowData);

    /// <summary>
    /// Cell value validator delegate, validate current cell value before <see cref="PropertyConfiguration.CellValueConverter"/>
    /// </summary>
    /// <param name="rowIndex">Row index of current cell in excel</param>
    /// <param name="columnIndex">Column index of current cell in excel</param>
    /// <param name="value">Value of current cell</param>
    /// <returns>Whether the value passes validation</returns>
    public delegate bool CellValueValidator(int rowIndex, int columnIndex, object value);

    /// <summary>
    /// Cell value converter delegate.
    /// </summary>
    /// <param name="rowIndex">Row index of current cell in excel</param>
    /// <param name="columnIndex">Column index of current cell in excel</param>
    /// <param name="value">Value of current cell</param>
    /// <returns>The converted value</returns>
    public delegate object CellValueConverter(int rowIndex, int columnIndex, object value);
}
