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
import SolidEdgeAssembly as SEAssembly


def main():
    application = SRI.Marshal.GetActiveObject("SolidEdge.Application")
    asm = application.ActiveDocument
    print("part: %s\n" % asm.Name)
    assert asm.Type == 3, "This macro only works on .asm"

    # Queries:
    objQueries = asm.Queries


    # -------------------------
    # "Hardware Plated Zinc"
    # -------------------------

    if True:
        query_name = "Hardware [PLATED.ZINC]"
        # Add the query here:
        zinc = objQueries.Add(query_name)
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
        print("[QUERY]: Created: {0:.<25}{1:.>10}".format(query_name, zinc.MatchesCount.ToString()))


    # -------------------------
    #  "Hardware [SS]"
    # -------------------------

    if True:
        query_name = "Hardware [SS]"
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
        print("[QUERY]: Created: {0:.<25}{1:.>10}".format(query_name, zinc.MatchesCount.ToString()))


    # -------------------------
    #  "Hardware [SS.304]"
    # -------------------------

    if True:
        query_name = "Hardware [SS.304]"

        ss304 = objQueries.Add(query_name)
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
        print("[QUERY]: Created: {0:.<25}{1:.>10}".format(query_name, ss304.MatchesCount.ToString()))


    # -------------------------
    #  "Hardware [SS.316]"
    # -------------------------

    if True:
        query_name = "Hardware [SS.316]"
        # Add the query here:
        ss316 = objQueries.Add(query_name)
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
        print("[QUERY]: Created: {0:.<25}{1:.>10}".format(query_name, ss316.MatchesCount.ToString()))


    # -------------------------
    #  "Hardware Imperial"
    # -------------------------

    if True:
        query_name = "Hardware INCH"
        imp = objQueries.Add(query_name)
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
            SEAssembly.QueryConditionConstants.seQueryConditionIsNot,
            "Metric Fastener",
        )
        print("[QUERY]: Created: {0:.<25}{1:.>10}".format(query_name, imp.MatchesCount.ToString()))


    # -------------------------
    #  "Hardware Metric"
    # -------------------------

    if True:
        query_name = "Hardware METRIC"
        metric = objQueries.Add(query_name)
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
            "Metric Fastener",
        )
        print("[QUERY]: Created: {0:.<25}{1:.>10}".format(query_name, metric.MatchesCount.ToString()))
# TODO: Add reference object
# TODO: add non released items


def confirmation(func):
    response = raw_input("""Create fasteners queries? (Press y/[Y] to proceed.)""")
    if response.lower() not in ["y", "yes"]:
        print("Process canceled")
        sys.exit()
    else:
        func()


if __name__ == "__main__":
    confirmation(main)