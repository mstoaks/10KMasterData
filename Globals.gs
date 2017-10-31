/*
File: Globals.gs
Author: Max Stoaks
Purpose: Global variables for related script files.

TODO:
1. Clean up global naming consistency (_ID vs no _ID)
*/

//Custom Attribute IDs - NOTE UPDATE IF 10K Custom Attribute IDs CHANGE
// i.e. if ever have to move to a new 10K db, these IDs might change
PROJ_PHASE_ID = 841;
PROJ_MGR_ID = 842;
PROJ_PROD_MGR = 843;
PROJ_STRAT_OBJ_ID = 1124;
PROJ_ARCH_PARTNER_ID = 1175;
PROJ_STATUS_ID = 1276;
PROJ_AC_DECK = 1447;
PROJ_PMO_DECK = 1448;
PROJ_DIRECTOR = 1515;
PROJ_GSB_PRIO = 1530;
PROJ_EFFORT = 1531;
PROJ_BENEFICIARY = 1532;
PROJ_VALUE = 1533;
PROJ_PGM_INDICATOR = 1534;
PROJ_PARENT_PGM = 1535;
PROJ_FOLDER = 1538;
PROJ_MICITI = 1543;
PROJ_PRMY_CONTACT = 1652;
PROJ_NOTES = 1745;
PROJ_NONCOMP_COST = 1926;
PROJ_HI_BENE_IMPACT = 1927;
PROJ_FORCE_PRIO = 1928;
PROJ_RECORD_MGR = 1946;
PROJ_OP_PRIO = 1953;



//to pre-fetch all resources so don't have to call API for the same resource repeatedly
allResources = null;

//to pre-fetch all projects
allProjects = null;

//params for api calls - all GET calls
params = null;
