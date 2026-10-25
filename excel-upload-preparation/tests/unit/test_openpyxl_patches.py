"""openpyxl 3.1 bugs the app patches (app.fast_xlsx.patch_openpyxl)."""
from openpyxl.pivot.cache import FieldGroup
from openpyxl.xml.functions import fromstring, tostring

import app.inventory  # noqa: F401 -- applies the patches

GROUPED = (
    '<fieldGroup xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" base="3">'
    '<discretePr count="3"><x v="0"/><x v="1"/><x v="0"/></discretePr>'
    '<groupItems count="2"><s v="A"/><s v="B"/></groupItems></fieldGroup>'
)


def test_a_grouped_pivot_field_reads_and_writes_back():
    # UC-5: "Nested.from_tree() missing 1 required positional argument: 'node'"
    fg = FieldGroup.from_tree(fromstring(GROUPED))
    assert [i.v for i in fg.discretePr] == [0, 1, 0]
    out = tostring(fg.to_tree()).decode()
    assert '<discretePr count="3"><x v="0" /><x v="1" /><x v="0" /></discretePr>' in out.replace('"/>', '" />')
