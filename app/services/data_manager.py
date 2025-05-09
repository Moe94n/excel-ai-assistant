"""
Data Manager for the Excel AI Assistant.
Handles loading, saving, and manipulating Excel and CSV data.
"""

import logging
import os
from typing import Tuple, List, Dict, Any, Optional

import numpy as np
import pandas as pd
from pandas.errors import ParserError, EmptyDataError


class DataManager:
    """Manager for handling Excel and CSV data operations"""

    def __init__(self):
        """Initialize the data manager"""
        self.df = None
        self.file_path = None
        self.file_type = None
        self.logger = logging.getLogger("DataManager")
        self.modified = False

    def load_file(self, file_path: str) -> Tuple[bool, str]:
        """
        Load data from Excel or CSV file

        Args:
            file_path: Path to the file to load

        Returns:
            Tuple of (success, error_message)
        """
        if not os.path.exists(file_path):
            return False, f"File not found: {file_path}"

        try:
            # Get file extension
            file_ext = os.path.splitext(file_path)[1].lower()

            # Load based on file extension
            if file_ext in ('.xlsx', '.xls', '.xlsm'):
                self.df = pd.read_excel(file_path)
                self.file_type = 'excel'
            elif file_ext == '.csv':
                # Try different encodings
                try:
                    self.df = pd.read_csv(file_path)
                except UnicodeDecodeError:
                    # Try with different encodings
                    for encoding in ['utf-8', 'latin1', 'iso-8859-1', 'cp1252']:
                        try:
                            self.df = pd.read_csv(file_path, encoding=encoding)
                            break
                        except UnicodeDecodeError:
                            continue
                    else:
                        return False, "Could not decode file with any common encoding"
            else:
                return False, f"Unsupported file format: {file_ext}"

            # Check if dataframe is valid
            if self.df is None or len(self.df) == 0:
                return False, "File contains no data"

            # Clean up column names (remove whitespace)
            self.df.columns = self.df.columns.str.strip()

            # Handle any NaN values gracefully
            self.df = self.df.replace({np.nan: None})

            self.file_path = file_path
            self.modified = False
            return True, ""

        except EmptyDataError:
            return False, "The file is empty"
        except ParserError:
            return False, "Error parsing the file. It may be corrupted."
        except PermissionError:
            return False, "Permission denied. The file may be in use by another program."
        except Exception as e:
            self.logger.error(f"Error loading file: {str(e)}")
            return False, f"Failed to load file: {str(e)}"

    def save_file(self, file_path: Optional[str] = None) -> Tuple[bool, str]:
        """
        Save data to Excel or CSV file

        Args:
            file_path: Path to save to (if None, use the current file_path)

        Returns:
            Tuple of (success, error_message)
        """
        if self.df is None:
            return False, "No data to save"

        # Use current path if not specified
        if file_path is None:
            if self.file_path is None:
                return False, "No file path specified"
            file_path = self.file_path

        try:
            # Get file extension
            file_ext = os.path.splitext(file_path)[1].lower()

            # Save based on file extension
            if file_ext in ('.xlsx', '.xls'):
                self.df.to_excel(file_path, index=False)
            elif file_ext == '.csv':
                self.df.to_csv(file_path, index=False)
            else:
                # Add default extension if none was specified
                if '.' not in os.path.basename(file_path):
                    file_path += '.xlsx'
                    self.df.to_excel(file_path, index=False)

            self.file_path = file_path
            self.modified = False
            return True, ""

        except PermissionError:
            return False, "Permission denied. The file may be in use by another program."
        except Exception as e:
            self.logger.error(f"Error saving file: {str(e)}")
            return False, f"Failed to save file: {str(e)}"

    def get_data(self) -> pd.DataFrame:
        """Get the current dataframe"""
        return self.df

    def get_meta_info(self) -> Dict[str, Any]:
        """Get metadata about the current dataframe"""
        if self.df is None:
            return {"loaded": False}

        return {
            "loaded": True,
            "file_path": self.file_path,
            "file_name": os.path.basename(self.file_path) if self.file_path else None,
            "file_type": self.file_type,
            "rows": len(self.df),
            "columns": len(self.df.columns),
            "column_names": list(self.df.columns),
            "modified": self.modified
        }

    def get_column_data_types(self) -> Dict[str, str]:
        """Get data types for each column"""
        if self.df is None:
            return {}

        types = {}
        for col in self.df.columns:
            # Get pandas dtype as string and simplify it
            dtype = str(self.df[col].dtype)

            if dtype.startswith('int'):
                types[col] = 'integer'
            elif dtype.startswith('float'):
                types[col] = 'float'
            elif dtype.startswith('datetime'):
                types[col] = 'datetime'
            elif dtype.startswith('bool'):
                types[col] = 'boolean'
            else:
                types[col] = 'text'

        return types

    def get_cell_value(self, row: int, col: str) -> Any:
        """Get the value of a specific cell"""
        if self.df is None or row < 0 or row >= len(self.df) or col not in self.df.columns:
            return None

        return self.df.loc[row, col]

    def update_cell(self, row: int, col: str, value: Any) -> bool:
        """Update the value of a specific cell"""
        if self.df is None or row < 0 or row >= len(self.df) or col not in self.df.columns:
            return False

        try:
            self.df.loc[row, col] = value
            self.modified = True
            return True
        except Exception as e:
            self.logger.error(f"Error updating cell: {str(e)}")
            return False

    def update_range(self, updates: List[Dict[str, Any]], auto_save: bool = False) -> Tuple[int, int]:
        """
        Update multiple cells at once and optionally auto-save

        Args:
            updates: List of dictionaries with row, col, and value keys
            auto_save: Whether to automatically save changes to the source file after updating

        Returns:
            Tuple of (updates_applied, failed_updates)
        """
        if self.df is None:
            return 0, len(updates)

        success_count = 0
        error_count = 0

        for update in updates:
            row = update.get('row')
            col = update.get('col')
            value = update.get('result')

            if row is None or col is None or value is None:
                error_count += 1
                continue

            if self.update_cell(row, col, value):
                success_count += 1
            else:
                error_count += 1

        if success_count > 0:
            self.modified = True
            
            # Auto-save if requested and we have a file path
            if auto_save and self.file_path and success_count > 0:
                self.logger.info(f"Auto-saving changes to {self.file_path}")
                try:
                    save_success, error_msg = self.save_file()
                    if not save_success:
                        self.logger.error(f"Auto-save failed: {error_msg}")
                    else:
                        self.logger.info("Auto-save successful")
                except Exception as e:
                    self.logger.error(f"Auto-save exception: {str(e)}")

        return success_count, error_count

    def get_range(self, start_row: int, end_row: int, columns: List[str], 
                context_columns: List[str] = None) -> List[Dict[str, Any]]:
        """
        Get a range of cells with optional context columns

        Args:
            start_row: Starting row index
            end_row: Ending row index (exclusive)
            columns: List of column names to process
            context_columns: Optional list of column names to include as context

        Returns:
            List of dictionaries with row, col, content, and optional context_data keys
        """
        if self.df is None:
            return []

        # Validate parameters
        if start_row < 0:
            start_row = 0

        if end_row > len(self.df):
            end_row = len(self.df)

        # Validate columns
        valid_columns = [col for col in columns if col in self.df.columns]
        if not valid_columns:
            return []
            
        # Validate context columns
        valid_context_columns = []
        if context_columns:
            valid_context_columns = [col for col in context_columns if col in self.df.columns and col not in valid_columns]

        # Build result
        result = []
        for row in range(start_row, end_row):
            for col in valid_columns:
                cell_data = {
                    'row': row,
                    'col': col,
                    'content': self.df.loc[row, col]
                }
                
                # Add context data from other columns if requested
                if valid_context_columns:
                    context_data = {}
                    # Add column headers as context
                    for ctx_col in valid_context_columns:
                        context_data[ctx_col] = self.df.loc[row, ctx_col]
                    
                    # Add column headers as additional context
                    context_data["headers"] = {col: col for col in valid_columns + valid_context_columns}
                    
                    cell_data['context_data'] = context_data
                
                result.append(cell_data)

        return result

    def get_data_summary(self) -> Dict[str, Any]:
        """Get a summary of the data"""
        if self.df is None:
            return {"summary": "No data loaded"}

        summary = {
            "rows": len(self.df),
            "columns": len(self.df.columns),
            "column_types": self.get_column_data_types(),
            "missing_values": self.df.isna().sum().to_dict(),
            "column_stats": {}
        }

        # Get basic stats for each column
        for col in self.df.columns:
            col_type = summary["column_types"][col]

            if col_type in ('integer', 'float'):
                # Numeric stats
                stats = {
                    "min": float(self.df[col].min()) if not pd.isna(self.df[col].min()) else None,
                    "max": float(self.df[col].max()) if not pd.isna(self.df[col].max()) else None,
                    "mean": float(self.df[col].mean()) if not pd.isna(self.df[col].mean()) else None,
                    "median": float(self.df[col].median()) if not pd.isna(self.df[col].median()) else None
                }
            elif col_type == 'datetime':
                # Date stats
                stats = {
                    "min": str(self.df[col].min()) if not pd.isna(self.df[col].min()) else None,
                    "max": str(self.df[col].max()) if not pd.isna(self.df[col].max()) else None
                }
            else:
                # Text stats
                unique_count = self.df[col].nunique()
                stats = {
                    "unique_values": unique_count,
                    "most_common": self.df[col].value_counts().head(5).to_dict() if unique_count < 100 else {}
                }

            # noinspection PyTypeChecker
            summary["column_stats"][col] = stats

        return summary

    # noinspection PyTypeChecker
    def analyze_column(self, column: str) -> Dict[str, Any]:
        """Analyze a specific column"""
        if self.df is None or column not in self.df.columns:
            return {"error": "Column not found"}

        col_data = self.df[column]
        col_type = self.get_column_data_types().get(column, 'unknown')

        # Start with basic information that works for all data types
        analysis = {
            "name": column,
            "type": col_type,
            "count": len(col_data),
            "missing": int(col_data.isna().sum()),
            "unique": int(col_data.nunique())
        }

        # Add type-specific stats with proper null handling
        if col_type in ('integer', 'float'):
            # Handle numeric statistics safely
            try:
                analysis["min"] = float(col_data.min()) if not pd.isna(col_data.min()) else None
            except (TypeError, ValueError):
                analysis["min"] = None

            try:
                analysis["max"] = float(col_data.max()) if not pd.isna(col_data.max()) else None
            except (TypeError, ValueError):
                analysis["max"] = None

            try:
                analysis["mean"] = float(col_data.mean()) if not pd.isna(col_data.mean()) else None
            except (TypeError, ValueError):
                analysis["mean"] = None

            try:
                analysis["median"] = float(col_data.median()) if not pd.isna(col_data.median()) else None
            except (TypeError, ValueError):
                analysis["median"] = None

            try:
                analysis["std"] = float(col_data.std()) if not pd.isna(col_data.std()) else None
            except (TypeError, ValueError):
                analysis["std"] = None

            # Calculate quartiles safely
            try:
                analysis["quartiles"] = {
                    "25%": float(col_data.quantile(0.25)) if not pd.isna(col_data.quantile(0.25)) else None,
                    "50%": float(col_data.quantile(0.5)) if not pd.isna(col_data.quantile(0.5)) else None,
                    "75%": float(col_data.quantile(0.75)) if not pd.isna(col_data.quantile(0.75)) else None
                }
            except (TypeError, ValueError):
                analysis["quartiles"] = {"25%": None, "50%": None, "75%": None}

        # Add histogram data for numeric columns
        if col_type in ('integer', 'float'):
            try:
                # Filter out null values before creating histogram
                valid_data = col_data.dropna()
                if len(valid_data) > 0:
                    hist_values, hist_edges = np.histogram(valid_data, bins=10)
                    analysis["histogram"] = {
                        "values": hist_values.tolist(),
                        "edges": hist_edges.tolist()
                    }
                else:
                    analysis["histogram"] = None
            except Exception:
                analysis["histogram"] = None

        # Add most common values for all types - with safe conversion to string
        try:
            value_counts = col_data.value_counts().head(10).to_dict()
            # Convert all keys to strings to avoid display issues
            analysis["most_common"] = {str(k): int(v) for k, v in value_counts.items()}
        except Exception:
            analysis["most_common"] = {}

        return analysis
