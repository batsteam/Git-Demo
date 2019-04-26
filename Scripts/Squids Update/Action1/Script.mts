Set oBrowser = Browser("Squids").Page("Home")

oBrowser.Link("Test Case & Defect").Click
oBrowser.Sync
oBrowser.WebEdit("TC Search").Set Datatable("TestCase",dtGlobalSheet)
oBrowser.Link("Search").Click
oBrowser.Sync
oBrowser.Link("TestCaseNbr").Click
oBrowser.Sync
oBrowser.Link("Next").Click
oBrowser.Sync
oBrowser.Link("OwnerDrop").Click
oBrowser.Link("Stacey, Ray (RXS0001)").Click
oBrowser.Link("PhaseDrop").Click
oBrowser.Link("Acceptance").Click
oBrowser.Sync
Wait(1)
oBrowser.Link("Save").Click
Wait(1)
oBrowser.Sync

 @@ hightlight id_;_Browser("Squids | Home").Page("Test Creation 3").Link("TCD RelatedTrackingWidget")_;_script infofile_;_ZIP::ssf23.xml_;_
