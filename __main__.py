# -*- coding: utf-8 -*-

"""
Query in SolidEdge:
==================

/// QueryScopeConstants Enumeration

seQueryScopeAllParts	    0
seQueryScopeHiddenParts	    2
seQueryScopeSelectedParts	3
seQueryScopeShownParts	    1


/// Query Property Constants

Member	                          Value
seQueryPropertyAuthor	            3
seQueryPropertyCategory	            6
seQueryPropertyComments	            8
seQueryPropertyCompany	            5
seQueryPropertyCustom	            15
seQueryPropertyCustomOccurrence	    16
seQueryPropertyDocumentNumber	    9
seQueryPropertyKeywords	            7
seQueryPropertyManager	            4
seQueryPropertyMaterial	            12
seQueryPropertyName	                0
seQueryPropertyProject	            11
seQueryPropertyReference	        14
seQueryPropertyRevisionNumber	    10
seQueryPropertyStatus	            13
seQueryPropertySubject	            2
seQueryPropertyTitle	            1


/// QueryConditionConstants Enumeration

seQueryConditionContains	0
seQueryConditionIs	        1
seQueryConditionIsNot	    2
"""

import sys
import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System")
clr.AddReference("System.Runtime.InteropServices")

import System
import System.Collections
import System.Runtime.InteropServices as SRI
from System import Console
# import SolidEdgeAssembly as SEAssembly
from SolidEdgeAssembly.QueryScopeConstants import  seQueryScopeAllParts
from SolidEdgeAssembly.QueryPropertyConstants import seQueryPropertyCategory
from SolidEdgeAssembly.QueryPropertyConstants import seQueryPropertyCustom
from SolidEdgeAssembly.QueryPropertyConstants import seQueryPropertyReference
from SolidEdgeAssembly.QueryConditionConstants import seQueryConditionContains
from SolidEdgeAssembly.QueryConditionConstants import seQueryConditionIsNot

def create_query(queries, query_name, criterias=None, sub=True):
    """Create a query in select tools panel.
    """

    # Add the query here:
    query = queries.Add(query_name)
    query.Scope = seQueryScopeAllParts
    query.SearchSubassemblies = sub

    # loop throught the criterias
    for criteria in criterias:
        query.AddCriteria(
            criteria[0],
            criteria[1],
            criteria[2],
            criteria[3],
    )
    print("[QUERY]: Created: {0:.<25}{1:.>25}".format(query_name, query.MatchesCount.ToString()))


def main():
    application = SRI.Marshal.GetActiveObject("SolidEdge.Application")
    asm = application.ActiveDocument
    print("part: %s\n" % asm.Name)
    assert asm.Type == 3, "This macro only works on .asm"

    fasteners = False

    # Hardware [PLATED.ZINC]
    create_query(
            asm.Queries,
            "Hardware [PLATED.ZINC]",
            [
                (seQueryPropertyCategory, "Category", seQueryConditionContains ,"HARDWARE"),
                (seQueryPropertyCustom, "DSC_A", seQueryConditionContains , "ZINC PLATED"),
            ]
    )

    # Hardware [SS]
    create_query(
            asm.Queries,
            "Hardware [SS]",
            [
                (seQueryPropertyCategory, "Category", seQueryConditionContains ,"HARDWARE"),
                (seQueryPropertyCustom, "DSC_F", seQueryConditionContains , "SS.3"),
            ]
    )
    # Hardware [SS.304]
    create_query(
            asm.Queries,
            "Hardware [SS.304]",
            [
                (seQueryPropertyCategory, "Category", seQueryConditionContains ,"HARDWARE"),
                (seQueryPropertyCustom, "DSC_F", seQueryConditionContains , "[SS.304]"),
            ]
    )
    # Hardware [SS.316]
    create_query(
            asm.Queries,
            "Hardware [SS.316]",
            [
                (seQueryPropertyCategory, "Category", seQueryConditionContains ,"HARDWARE"),
                (seQueryPropertyCustom, "DSC_F", seQueryConditionContains , "[SS.316]"),
            ]
    )
    # "Hardware INCH"
    create_query(
            asm.Queries,
            "Hardware INCH",
            [
                (seQueryPropertyCategory, "Category", seQueryConditionContains ,"HARDWARE"),
                (seQueryPropertyCustom, "JDEPRP1", seQueryConditionIsNot , "Metric Fastener"),
            ]
    )
    # "Hardware METRIC"
    create_query(
            asm.Queries,
            "Hardware METRIC",
            [
                (seQueryPropertyCategory, "Category", seQueryConditionContains ,"HARDWARE"),
                (seQueryPropertyCustom, "JDEPRP1", seQueryConditionContains , "Metric Fastener"),
            ]
    )



    # "Reference"
    create_query(
            asm.Queries,
            "Reference",
            [
                (seQueryPropertyReference, "Reference", seQueryConditionContains ,"HARDWARE"),
            ]
    )
    # "Reference"
    create_query(
            asm.Queries,
            "Reference",
            [
                (seQueryPropertyReference, "Reference", seQueryConditionContains ,"HARDWARE"),
            ]
    )
    # "Reference"
    create_query(
            asm.Queries,
            "Reference",
            [
                (seQueryPropertyReference, "Reference", seQueryConditionContains ,"HARDWARE"),
            ]
    )



# TODO: Add reference object
# TODO: add non released items


def confirmation(func):
    response = raw_input("""Create fasteners/reference queries? (Press y/[Y] to proceed.)""")
    if response.lower() not in ["y", "yes"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


if __name__ == "__main__":
    confirmation(main)
