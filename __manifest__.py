{
	'name': 'FINISH FET PLAN Module',
    'summary' : "FINISH FET PLAN",
	'description' : """FINISH FET PLAN""",
	'author' : "ISGEC IT",
	'license' : "AGPL-3",
	'website' : "www.isgec.com",
	'category' : 'Uncategorized',
	'version' : '12.0.1.0.0',
	'depends' : ['base'],
	'data' : [
		     'security/groups.xml',
		     'views/FinishFetPlanModule_JobRoutingTable.xml',
		     'views/FinishFetPlanModule_ManpowerTable.xml',
		     'views/FinishFetPlanModule_ItemPlanTable.xml',
		     'views/FinishFetPlanModule_ItemPlanHeaderTable.xml',
		     'wizard/FinishFetPlanModule_FinishFetPlanReport_view.xml',
		     'wizard/FinishFetPlanModule_ActualItemReport_view.xml',
		     'views/FinishFetPlanModule_ActualItemPlanTable.xml',
	         'security/ir.model.access.csv',

	         ],	
}
