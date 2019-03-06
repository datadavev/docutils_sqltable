"""
SQLTable is an extension to docutils reStructuredText processor which adds
the ability to pull table content from an SQL data source, a CSV file, or
an Excel spread sheet.

This project was originally started within the DataONE project (https://dataone.org)

Example for use in Sphinx, add to conf.py after ``import sys, os``::

  #Import the SQLTable directive RST extension
  from sqltable import SQLTable
  from docutils.parsers.rst import directives
  directives.register_directive('sql-table', SQLTable)
"""

__docformat__ = "reStructuredText"

import sys
import os.path
import csv

from docutils import io, nodes, statemachine, utils
from docutils.utils import SystemMessagePropagation
from docutils.parsers.rst import Directive
from docutils.parsers.rst import directives
from docutils.parsers.rst.directives import tables


class SQLTable(tables.Table):
    """
    """

    required_arguments = 3
    optional_arguments = 5
    final_argument_whitespace = True
    has_content = True
    option_spec = {
        "header": directives.unchanged,
        "widths": directives.positive_int_list,
        "encoding": directives.encoding,
        "stub-columns": directives.nonnegative_int,
        "class": directives.class_option,
        "driver": directives.unchanged_required,
        "source": directives.path,
        "sql": directives.unchanged_required,
    }

    def check_requirements(self):
        pass

    def run(self):
        try:
            self.check_requirements()
            stub_columns = self.options.get("stub-columns", 0)
            title, messages = self.make_title()
            table_body, max_cols = self.get_sql_data()
            table_head = self.process_header_option()
            col_widths = self.get_column_widths(max_cols)
            # hack until figure out change in docutils
            try:
                a = int(col_widths[0])
            except ValueError as e:
                col_widths = self.get_column_widths(max_cols)[1]

            self.check_table_dimensions(table_body, 0, stub_columns)
            self.extend_short_rows_with_empty_cells(max_cols, (table_head, table_body))
        except SystemMessagePropagation as detail:
            return [detail.args[0]]
        except csv.Error as detail:
            error = self.state_machine.reporter.error(
                'Error with CSV data in "%s" directive:\n%s' % (self.name, detail),
                nodes.literal_block(self.block_text, self.block_text),
                line=self.lineno,
            )
            return [error]
        table = (col_widths, table_head, table_body)
        table_node = self.state.build_table(table, self.content_offset, stub_columns)
        table_node["classes"] += self.options.get("class", [])
        if title:
            table_node.insert(0, title)
        return [table_node] + messages

    def process_header_option(self):
        """Returns table_head
        """
        res = []
        colnames = self.options.get("header", "")
        for col in colnames.split(","):
            res.append([0, 0, 0, statemachine.StringList(col.strip().splitlines())])
        return [res]

    def get_sql_data(self):
        """Returns rows, max_cols
        """
        # Load the specified driver and get a connection to the database
        driver = self.options.get("driver", "xlsx")
        source_dir = os.path.dirname(
            os.path.abspath(self.state.document.current_source))
        if driver == "xlsx":
            # Load content to an in-memory SQLite database.
            # Yeah, it's ugly, but it works pretty well actually.
            from .xls2sql import Xls2Sql
            import sqlite3

            dbconn = sqlite3.connect(":memory:")
            loader = Xls2Sql(dbconn)
            source_path = os.path.normpath(os.path.join(source_dir, self.options.get("source", "data.xlsx")))
            source_path = utils.relative_path(None, source_path)
            loader.load(os.path.abspath(source_path))
        else:
            #default_src = 'database="data.xlsx",driver="xslx"'
            default_src = 'database="csv-data",driver="csv"'
            exec("import %s as DBDRV" % self.options.get("driver", "SnakeSQL"))
            cnstr = "dbconn = DBDRV.connect(%s)" % self.options.get(
                "source", default_src
            )
            exec(cnstr)
        cursor = dbconn.cursor()
        SQL = str(self.options.get("sql"))
        res = cursor.execute(SQL)
        rows = []
        max_cols = 0
        row = cursor.fetchone()
        while row is not None:
            row_data = []
            for cell in row:
                cell_text = str(cell)
                cell_data = (0, 0, 0, statemachine.StringList(cell_text.splitlines()))
                row_data.append(cell_data)
            rows.append(row_data)
            max_cols = max(max_cols, len(row_data))
            row = cursor.fetchone()
        dbconn.close()
        return rows, max_cols
