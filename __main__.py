# -*- coding: utf-8 -*-

"""
Query in SolidEdge:
==================
Description: Query Property Constants

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
import SolidEdgeAssembly as SEAssembly


def isQueryExist(queryName, queries):
    if queries.Item(queryName):
        return True
    False

def main():
    application = SRI.Marshal.GetActiveObject("SolidEdge.Application")
    asm = application.ActiveDocument
    print("part: %s\n" % asm.Name)
    assert asm.Type == 3, "This macro only works on .asm"

    # ActiveSelectSet
    selectSet = application.ActiveSelectSet
    selectSet.RemoveAll()

    # Queries:
    objQueries = asm.Queries
    # TODO: check if query exist already then skip
    objQueries.Item("Hardware")
    try:
        print(objQueries.Count)
    except Exception:
        return False

    

"""
    # "Hardware Plated Zinc"
    # =====================
    if True:
        # Add the query here:
        zinc = objQueries.Add("Plated zinc")
        zinc.Scope = SEAssembly.QueryScopeConstants.seQueryScopeAllParts
        zinc.SearchSubassemblies = False

        # Add Criteria to above query
        zinc.AddCriteria(
            SEAssembly.QueryPropertyConstants.seQueryPropertyCategory,
            "Category",
            SEAssembly.QueryConditionConstants.seQueryConditionContains,
            "HARDWARE",
        )
        # Add a second criteria
        zinc.AddCriteria(
            SEAssembly.QueryPropertyConstants.seQueryPropertyCustom,
            "DSC_A",
            SEAssembly.QueryConditionConstants.seQueryConditionContains,
            "ZINC PLATED",
        )
        print("[QUERY] - 'Hardware Plated Zinc' created ->\t\t qty: %s" % zinc.MatchesCount.ToString())

    # -------------------------
    #  "Hardware [SS]"
    # -------------------------

    # Add the query here:
    ss = objQueries.Add("Hardware [SS]")
    ss.Scope = SEAssembly.QueryScopeConstants.seQueryScopeAllParts
    ss.SearchSubassemblies = False

    # Add Criteria to above query
    ss.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCategory,
        "Category",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "HARDWARE",
    )
    # Add Criteria to above query
    ss.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCustom,
        "DSC_F",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "SS.3",
    )
    print("[QUERY]: Hardware [SS] qty: .... %s" % ss.MatchesCount.ToString())

    # -------------------------
    #  "Hardware [SS.304]"
    # -------------------------

    # Add the query here:
    # TODO: name of the query contained in a variable
    ss304 = objQueries.Add("Hardware [SS.304]")
    ss304.Scope = SEAssembly.QueryScopeConstants.seQueryScopeAllParts
    ss304.SearchSubassemblies = False

    ss304.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCategory,
        "Category",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "HARDWARE",
    )
    # Add Criteria to above query
    ss304.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCustom,
        "DSC_F",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "[SS.304]",
    )
    print("[QUERY] - 'Hardware [SS.304]' ->\t\t qty: %s" % ss304.MatchesCount.ToString()) # TODO: format the output with {:>.}

    # TODO: review the order of the queries in solidedge.
    # -------------------------
    #  "Hardware [SS.316]"
    # -------------------------

    # Add the query here:
    ss316 = objQueries.Add("Hardware [SS.316]")
    ss316.Scope = SEAssembly.QueryScopeConstants.seQueryScopeAllParts
    ss316.SearchSubassemblies = False

    ss316.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCategory,
        "Category",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "HARDWARE",
    )
    # Add Criteria to above query
    ss316.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCustom,
        "DSC_F",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "[SS.316]",
    )
    print("[QUERY] - 'Hardware [SS.316]' ->\t\t qty: %s" % ss316.MatchesCount.ToString())

    # -------------------------
    #  "Hardware Imperial"
    # -------------------------

    # Add the query here:
    # if not objQueries.Item("Hardware imperial"):
    imp = objQueries.Add("Hardware imperial")
    imp.Scope = SEAssembly.QueryScopeConstants.seQueryScopeAllParts
    imp.SearchSubassemblies = False

    # Add Criteria to above query
    imp.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCategory,
        "Category",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "HARDWARE",
    )
    # Add Criteria to above query
    imp.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCustom,
        "JDEPRP1",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "Metric Fastener",
    )
    print("[QUERY] - 'Hardware imperial' ->\t\t qty: %s" % imp.MatchesCount.ToString())

    # -------------------------
    #  "Hardware Metric"
    # -------------------------

    # Add the query here:
    # if not objQueries.Item("Hardware metric"):
    metric = objQueries.Add("Hardware metric")
    metric.Scope = SEAssembly.QueryScopeConstants.seQueryScopeAllParts
    metric.SearchSubassemblies = False

    # Add Criteria to above query
    metric.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCategory,
        "Category",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "HARDWARE",
    )
    # Add Criteria to above query
    metric.AddCriteria(
        SEAssembly.QueryPropertyConstants.seQueryPropertyCustom,
        "JDEPRP1",
        SEAssembly.QueryConditionConstants.seQueryConditionContains,
        "Inch Fastener",
    )
    print("[QUERY] - 'Hardware metric' ->\t\t qty: %s" % metric.MatchesCount.ToString())

# TODO: Add reference object
# TODO: add non released items

"""

def confirmation(func):
    response = raw_input("""Create fasteners queries? (Press y/[Y] to proceed.)""")
    if response.lower() not in ["y", "yes", "oui"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


if __name__ == "__main__":
    confirmation(main)