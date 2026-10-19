Attribute VB_Name = "modReportPageMap"
Option Explicit
Public Function ReportPageForField(ByVal kind As String, ByVal key As String) As Long
Select Case kind & ":" & key
Case "EVAL:PeriodStart": ReportPageForField = 0
Case "EVAL:PeriodEnd": ReportPageForField = 0
Case "EVAL:OfficialComments": ReportPageForField = 1
Case "EVAL:OfficialName": ReportPageForField = 0
Case "EVAL:RankRate": ReportPageForField = 0
Case "EVAL:Designator": ReportPageForField = 0
Case "EVAL:UnitUIC": ReportPageForField = 0
Case "EVAL:ShipStation": ReportPageForField = 0
Case "EVAL:Trait33": ReportPageForField = 0
Case "EVAL:Trait34": ReportPageForField = 0
Case "EVAL:Trait35": ReportPageForField = 0
Case "EVAL:Trait36": ReportPageForField = 0
Case "EVAL:Trait37": ReportPageForField = 0
Case "EVAL:Trait38": ReportPageForField = 0
Case "EVAL:Trait39": ReportPageForField = 0
Case "EVAL:ObservationStatus": ReportPageForField = 0
Case "EVAL:PromotionRecommendation": ReportPageForField = 1
Case "EVAL:PromotionStatus": ReportPageForField = 0
Case "EVAL:ReportingSeniorName": ReportPageForField = 0
Case "EVAL:ReportingSeniorGrade": ReportPageForField = 0
Case "EVAL:ReportingSeniorDesignator": ReportPageForField = 0
Case "EVAL:ReportingSeniorTitle": ReportPageForField = 0
Case "EVAL:ReportingSeniorUIC": ReportPageForField = 0
Case "EVAL:OccasionPeriodic": ReportPageForField = 0
Case "EVAL:OccasionDetachment": ReportPageForField = 0
Case "EVAL:OccasionFrocking": ReportPageForField = 0
Case "EVAL:OccasionSpecial": ReportPageForField = 0
Case "EVAL:ReportRegular": ReportPageForField = 0
Case "EVAL:ReportConcurrent": ReportPageForField = 0
Case "EVAL:ReportOperationalCommander": ReportPageForField = 0
Case "EVAL:PrimaryDuty": ReportPageForField = 0
Case "EVAL:DutyDescription": ReportPageForField = 0
Case "EVAL:MemberSSN": ReportPageForField = 0
Case "EVAL:ComponentActive": ReportPageForField = 0
Case "EVAL:ComponentTAR": ReportPageForField = 0
Case "EVAL:ComponentInactive": ReportPageForField = 0
Case "EVAL:ComponentATADSW": ReportPageForField = 0
Case "EVAL:DateReported": ReportPageForField = 0
Case "EVAL:CounselingDate": ReportPageForField = 0
Case "EVAL:CounselorName": ReportPageForField = 0
Case "EVAL:ReadinessSummary": ReportPageForField = 0
Case "EVAL:ReadinessDate": ReportPageForField = 0
Case "EVAL:BilletSubcategory": ReportPageForField = 0
Case "EVAL:AssignmentRecommendationA": ReportPageForField = 1
Case "EVAL:AssignmentRecommendationB": ReportPageForField = 1
Case "EVAL:RetentionYes": ReportPageForField = 1
Case "EVAL:RetentionNo": ReportPageForField = 1
Case "EVAL:StatementYes": ReportPageForField = 1
Case "EVAL:StatementNo": ReportPageForField = 1
Case "EVAL:ReportingSeniorSSN": ReportPageForField = 0
Case "EVAL:ReportingSeniorAddress1": ReportPageForField = 1
Case "EVAL:ReportingSeniorAddress2": ReportPageForField = 1
Case "EVAL:ReportingSeniorCity": ReportPageForField = 1
Case "EVAL:ReportingSeniorState": ReportPageForField = 1
Case "EVAL:ReportingSeniorZip": ReportPageForField = 1
Case "EVAL:ReportingSeniorPhone": ReportPageForField = 1
Case "EVAL:ReportingSeniorDSN": ReportPageForField = 1
Case "EVAL:RaterName": ReportPageForField = 1
Case "EVAL:RaterDate": ReportPageForField = 1
Case "EVAL:SeniorRaterName": ReportPageForField = 1
Case "EVAL:SeniorRaterDate": ReportPageForField = 1
Case "EVAL:CommandEmployment": ReportPageForField = 0
Case "CHIEFEVAL:PeriodStart": ReportPageForField = 0
Case "CHIEFEVAL:PeriodEnd": ReportPageForField = 0
Case "CHIEFEVAL:OfficialComments": ReportPageForField = 1
Case "CHIEFEVAL:OfficialName": ReportPageForField = 0
Case "CHIEFEVAL:RankRate": ReportPageForField = 0
Case "CHIEFEVAL:Designator": ReportPageForField = 0
Case "CHIEFEVAL:UnitUIC": ReportPageForField = 0
Case "CHIEFEVAL:ShipStation": ReportPageForField = 0
Case "CHIEFEVAL:Trait33": ReportPageForField = 0
Case "CHIEFEVAL:Trait34": ReportPageForField = 0
Case "CHIEFEVAL:Trait35": ReportPageForField = 0
Case "CHIEFEVAL:Trait36": ReportPageForField = 0
Case "CHIEFEVAL:Trait37": ReportPageForField = 0
Case "CHIEFEVAL:Trait38": ReportPageForField = 0
Case "CHIEFEVAL:Trait39": ReportPageForField = 0
Case "CHIEFEVAL:ObservationStatus": ReportPageForField = 0
Case "CHIEFEVAL:PromotionRecommendation": ReportPageForField = 1
Case "CHIEFEVAL:PromotionStatus": ReportPageForField = 0
Case "CHIEFEVAL:ReportingSeniorName": ReportPageForField = 0
Case "CHIEFEVAL:ReportingSeniorGrade": ReportPageForField = 0
Case "CHIEFEVAL:ReportingSeniorDesignator": ReportPageForField = 0
Case "CHIEFEVAL:ReportingSeniorTitle": ReportPageForField = 0
Case "CHIEFEVAL:ReportingSeniorUIC": ReportPageForField = 0
Case "CHIEFEVAL:OccasionPeriodic": ReportPageForField = 0
Case "CHIEFEVAL:OccasionDetachment": ReportPageForField = 0
Case "CHIEFEVAL:OccasionFrocking": ReportPageForField = 0
Case "CHIEFEVAL:OccasionSpecial": ReportPageForField = 0
Case "CHIEFEVAL:ReportRegular": ReportPageForField = 0
Case "CHIEFEVAL:ReportConcurrent": ReportPageForField = 0
Case "CHIEFEVAL:ReportOperationalCommander": ReportPageForField = 0
Case "CHIEFEVAL:PrimaryDuty": ReportPageForField = 0
Case "CHIEFEVAL:DutyDescription": ReportPageForField = 0
Case "CHIEFEVAL:MemberSSN": ReportPageForField = 0
Case "CHIEFEVAL:ComponentActive": ReportPageForField = 0
Case "CHIEFEVAL:ComponentTAR": ReportPageForField = 0
Case "CHIEFEVAL:ComponentInactive": ReportPageForField = 0
Case "CHIEFEVAL:ComponentATADSW": ReportPageForField = 0
Case "CHIEFEVAL:DateReported": ReportPageForField = 0
Case "CHIEFEVAL:CounselingDate": ReportPageForField = 0
Case "CHIEFEVAL:CounselorName": ReportPageForField = 0
Case "CHIEFEVAL:ReadinessSummary": ReportPageForField = 0
Case "CHIEFEVAL:ReadinessDate": ReportPageForField = 0
Case "CHIEFEVAL:BilletSubcategory": ReportPageForField = 0
Case "CHIEFEVAL:AssignmentRecommendationA": ReportPageForField = 1
Case "CHIEFEVAL:AssignmentRecommendationB": ReportPageForField = 1
Case "CHIEFEVAL:RetentionYes": ReportPageForField = 1
Case "CHIEFEVAL:RetentionNo": ReportPageForField = 1
Case "CHIEFEVAL:StatementYes": ReportPageForField = 1
Case "CHIEFEVAL:StatementNo": ReportPageForField = 1
Case "CHIEFEVAL:ReportingSeniorSSN": ReportPageForField = 0
Case "CHIEFEVAL:ReportingSeniorAddress1": ReportPageForField = 1
Case "CHIEFEVAL:ReportingSeniorAddress2": ReportPageForField = 1
Case "CHIEFEVAL:ReportingSeniorCity": ReportPageForField = 1
Case "CHIEFEVAL:ReportingSeniorState": ReportPageForField = 1
Case "CHIEFEVAL:ReportingSeniorZip": ReportPageForField = 1
Case "CHIEFEVAL:ReportingSeniorPhone": ReportPageForField = 1
Case "CHIEFEVAL:ReportingSeniorDSN": ReportPageForField = 1
Case "CHIEFEVAL:RaterName": ReportPageForField = 1
Case "CHIEFEVAL:RaterDate": ReportPageForField = 1
Case "CHIEFEVAL:SeniorRaterName": ReportPageForField = 1
Case "CHIEFEVAL:SeniorRaterDate": ReportPageForField = 1
Case "CHIEFEVAL:CommandEmployment": ReportPageForField = 0
Case "FITREP:PeriodStart": ReportPageForField = 0
Case "FITREP:PeriodEnd": ReportPageForField = 0
Case "FITREP:OfficialComments": ReportPageForField = 1
Case "FITREP:OfficialName": ReportPageForField = 0
Case "FITREP:RankRate": ReportPageForField = 0
Case "FITREP:Designator": ReportPageForField = 0
Case "FITREP:UnitUIC": ReportPageForField = 0
Case "FITREP:ShipStation": ReportPageForField = 0
Case "FITREP:Trait33": ReportPageForField = 0
Case "FITREP:Trait34": ReportPageForField = 0
Case "FITREP:Trait35": ReportPageForField = 0
Case "FITREP:Trait36": ReportPageForField = 0
Case "FITREP:Trait37": ReportPageForField = 0
Case "FITREP:Trait38": ReportPageForField = 0
Case "FITREP:Trait39": ReportPageForField = 0
Case "FITREP:ObservationStatus": ReportPageForField = 0
Case "FITREP:PromotionRecommendation": ReportPageForField = 1
Case "FITREP:PromotionStatus": ReportPageForField = 0
Case "FITREP:ReportingSeniorName": ReportPageForField = 0
Case "FITREP:ReportingSeniorGrade": ReportPageForField = 0
Case "FITREP:ReportingSeniorDesignator": ReportPageForField = 0
Case "FITREP:ReportingSeniorTitle": ReportPageForField = 0
Case "FITREP:ReportingSeniorUIC": ReportPageForField = 0
Case "FITREP:OccasionPeriodic": ReportPageForField = 0
Case "FITREP:OccasionDetachment": ReportPageForField = 0
Case "FITREP:OccasionFrocking": ReportPageForField = 0
Case "FITREP:OccasionSpecial": ReportPageForField = 0
Case "FITREP:ReportRegular": ReportPageForField = 0
Case "FITREP:ReportConcurrent": ReportPageForField = 0
Case "FITREP:ReportOperationalCommander": ReportPageForField = 0
Case "FITREP:PrimaryDuty": ReportPageForField = 0
Case "FITREP:DutyDescription": ReportPageForField = 0
Case "FITREP:MemberSSN": ReportPageForField = 0
Case "FITREP:ComponentActive": ReportPageForField = 0
Case "FITREP:ComponentTAR": ReportPageForField = 0
Case "FITREP:ComponentInactive": ReportPageForField = 0
Case "FITREP:ComponentATADSW": ReportPageForField = 0
Case "FITREP:DateReported": ReportPageForField = 0
Case "FITREP:CounselingDate": ReportPageForField = 0
Case "FITREP:CounselorName": ReportPageForField = 0
Case "FITREP:ReadinessSummary": ReportPageForField = 0
Case "FITREP:ReadinessDate": ReportPageForField = 0
Case "FITREP:BilletSubcategory": ReportPageForField = 0
Case "FITREP:AssignmentRecommendationA": ReportPageForField = 1
Case "FITREP:AssignmentRecommendationB": ReportPageForField = 1
Case "FITREP:RetentionYes": ReportPageForField = 1
Case "FITREP:RetentionNo": ReportPageForField = 1
Case "FITREP:StatementYes": ReportPageForField = 1
Case "FITREP:StatementNo": ReportPageForField = 1
Case "FITREP:ReportingSeniorSSN": ReportPageForField = 0
Case "FITREP:ReportingSeniorAddress1": ReportPageForField = 1
Case "FITREP:ReportingSeniorAddress2": ReportPageForField = 1
Case "FITREP:ReportingSeniorCity": ReportPageForField = 1
Case "FITREP:ReportingSeniorState": ReportPageForField = 1
Case "FITREP:ReportingSeniorZip": ReportPageForField = 1
Case "FITREP:ReportingSeniorPhone": ReportPageForField = 1
Case "FITREP:ReportingSeniorDSN": ReportPageForField = 1
Case "FITREP:RaterName": ReportPageForField = 1
Case "FITREP:RaterDate": ReportPageForField = 1
Case "FITREP:SeniorRaterName": ReportPageForField = 1
Case "FITREP:SeniorRaterDate": ReportPageForField = 1
Case "FITREP:CommandEmployment": ReportPageForField = 0
End Select
End Function
Public Function ReportPageRows(ByVal kind As String, ByVal page As Long) As Collection
Dim result As New Collection
Select Case kind & ":" & page
Case "EVAL:0"
result.Add "OfficialName|180|180|5400|360"
result.Add "RankRate|5740|180|1650|360"
result.Add "Designator|7550|180|1650|360"
result.Add "MemberSSN|9360|180|2420|360"
result.Add "ComponentActive|180|1110|2192|360"
result.Add "ComponentTAR|2532|1110|2192|360"
result.Add "ComponentInactive|4884|1110|2192|360"
result.Add "ComponentATADSW|7236|1110|2192|360"
result.Add "UnitUIC|9588|1110|2192|360"
result.Add "ShipStation|180|2040|6100|360"
result.Add "PromotionStatus|6440|2040|2500|360"
result.Add "DateReported|9100|2040|2680|360"
result.Add "OccasionPeriodic|180|2970|2780|360"
result.Add "OccasionDetachment|3120|2970|2780|360"
result.Add "OccasionFrocking|6060|2970|2780|360"
result.Add "OccasionSpecial|9000|2970|2780|360"
result.Add "PeriodStart|180|3620|1900|360"
result.Add "PeriodEnd|2240|3620|1900|360"
result.Add "ObservationStatus|4300|3620|2400|360"
result.Add "ReportRegular|6860|3620|1500|360"
result.Add "ReportConcurrent|8520|3620|1600|360"
result.Add "ReportOperationalCommander|10280|3620|1500|360"
result.Add "ReadinessSummary|180|4550|3760|360"
result.Add "ReadinessDate|4100|4550|3760|360"
result.Add "BilletSubcategory|8020|4550|3760|360"
result.Add "ReportingSeniorName|180|5480|1800|360"
result.Add "ReportingSeniorGrade|2140|5480|1800|360"
result.Add "ReportingSeniorDesignator|4100|5480|1800|360"
result.Add "ReportingSeniorTitle|6060|5480|1800|360"
result.Add "ReportingSeniorUIC|8020|5480|1800|360"
result.Add "ReportingSeniorSSN|9980|5480|1800|360"
result.Add "CommandEmployment|180|6410|11600|950"
result.Add "PrimaryDuty|180|7930|5720|360"
result.Add "DutyDescription|6060|7930|5720|950"
result.Add "CounselingDate|180|9450|5720|360"
result.Add "CounselorName|6060|9450|5720|360"
result.Add "Trait33|180|10380|5000|360"
result.Add "Trait34|180|11000|5000|360"
result.Add "Trait35|180|11620|5000|360"
result.Add "Trait36|180|12240|5000|360"
result.Add "Trait37|180|12860|5000|360"
result.Add "Trait38|180|13480|5000|360"
result.Add "Trait39|180|14100|5000|360"
result.Add "@guidance|5550|10380|6000|4260"
Case "EVAL:1"
result.Add "AssignmentRecommendationA|180|180|3760|360"
result.Add "AssignmentRecommendationB|4100|180|3760|360"
result.Add "RaterName|8020|180|3760|360"
result.Add "RaterDate|180|1110|11600|360"
result.Add "OfficialComments|180|2040|11600|3924"
result.Add "PromotionRecommendation|180|7514|3760|360"
result.Add "RetentionYes|4100|7514|3760|360"
result.Add "RetentionNo|8020|7514|3760|360"
result.Add "ReportingSeniorAddress1|180|8444|3760|360"
result.Add "ReportingSeniorAddress2|4100|8444|3760|360"
result.Add "ReportingSeniorCity|8020|8444|3760|360"
result.Add "ReportingSeniorState|180|9374|3760|360"
result.Add "ReportingSeniorZip|4100|9374|3760|360"
result.Add "ReportingSeniorPhone|8020|9374|3760|360"
result.Add "ReportingSeniorDSN|180|10304|3760|360"
result.Add "SeniorRaterName|4100|10304|3760|360"
result.Add "SeniorRaterDate|8020|10304|3760|360"
result.Add "StatementYes|180|11234|5720|360"
result.Add "StatementNo|6060|11234|5720|360"
Case "CHIEFEVAL:0"
result.Add "OfficialName|180|180|5400|360"
result.Add "RankRate|5740|180|1650|360"
result.Add "Designator|7550|180|1650|360"
result.Add "MemberSSN|9360|180|2420|360"
result.Add "ComponentActive|180|1110|2192|360"
result.Add "ComponentTAR|2532|1110|2192|360"
result.Add "ComponentInactive|4884|1110|2192|360"
result.Add "ComponentATADSW|7236|1110|2192|360"
result.Add "UnitUIC|9588|1110|2192|360"
result.Add "ShipStation|180|2040|6100|360"
result.Add "PromotionStatus|6440|2040|2500|360"
result.Add "DateReported|9100|2040|2680|360"
result.Add "OccasionPeriodic|180|2970|2780|360"
result.Add "OccasionDetachment|3120|2970|2780|360"
result.Add "OccasionFrocking|6060|2970|2780|360"
result.Add "OccasionSpecial|9000|2970|2780|360"
result.Add "PeriodStart|180|3620|1900|360"
result.Add "PeriodEnd|2240|3620|1900|360"
result.Add "ObservationStatus|4300|3620|2400|360"
result.Add "ReportRegular|6860|3620|1500|360"
result.Add "ReportConcurrent|8520|3620|1600|360"
result.Add "ReportOperationalCommander|10280|3620|1500|360"
result.Add "ReadinessSummary|180|4550|3760|360"
result.Add "ReadinessDate|4100|4550|3760|360"
result.Add "BilletSubcategory|8020|4550|3760|360"
result.Add "ReportingSeniorName|180|5480|1800|360"
result.Add "ReportingSeniorGrade|2140|5480|1800|360"
result.Add "ReportingSeniorDesignator|4100|5480|1800|360"
result.Add "ReportingSeniorTitle|6060|5480|1800|360"
result.Add "ReportingSeniorUIC|8020|5480|1800|360"
result.Add "ReportingSeniorSSN|9980|5480|1800|360"
result.Add "CommandEmployment|180|6410|11600|950"
result.Add "PrimaryDuty|180|7930|5720|360"
result.Add "DutyDescription|6060|7930|5720|950"
result.Add "CounselingDate|180|9450|5720|360"
result.Add "CounselorName|6060|9450|5720|360"
result.Add "Trait33|180|10380|5000|360"
result.Add "Trait34|180|11000|5000|360"
result.Add "Trait35|180|11620|5000|360"
result.Add "Trait36|180|12240|5000|360"
result.Add "Trait37|180|12860|5000|360"
result.Add "Trait38|180|13480|5000|360"
result.Add "Trait39|180|14100|5000|360"
result.Add "@guidance|5550|10380|6000|4260"
Case "CHIEFEVAL:1"
result.Add "OfficialComments|180|180|11600|1836"
result.Add "PromotionRecommendation|180|3566|3760|360"
result.Add "RaterName|4100|3566|3760|360"
result.Add "RaterDate|8020|3566|3760|360"
result.Add "RetentionYes|180|4496|3760|360"
result.Add "RetentionNo|4100|4496|3760|360"
result.Add "StatementYes|8020|4496|3760|360"
result.Add "StatementNo|180|5146|3760|360"
result.Add "SeniorRaterName|4100|5146|3760|360"
result.Add "SeniorRaterDate|8020|5146|3760|360"
result.Add "ReportingSeniorAddress1|180|6076|3760|360"
result.Add "ReportingSeniorAddress2|4100|6076|3760|360"
result.Add "ReportingSeniorCity|8020|6076|3760|360"
result.Add "ReportingSeniorState|180|7006|3760|360"
result.Add "ReportingSeniorZip|4100|7006|3760|360"
result.Add "ReportingSeniorPhone|8020|7006|3760|360"
result.Add "ReportingSeniorDSN|180|7936|3760|360"
result.Add "AssignmentRecommendationA|4100|7936|3760|360"
result.Add "AssignmentRecommendationB|8020|7936|3760|360"
Case "FITREP:0"
result.Add "OfficialName|180|180|5400|360"
result.Add "RankRate|5740|180|1650|360"
result.Add "Designator|7550|180|1650|360"
result.Add "MemberSSN|9360|180|2420|360"
result.Add "ComponentActive|180|1110|2192|360"
result.Add "ComponentTAR|2532|1110|2192|360"
result.Add "ComponentInactive|4884|1110|2192|360"
result.Add "ComponentATADSW|7236|1110|2192|360"
result.Add "UnitUIC|9588|1110|2192|360"
result.Add "ShipStation|180|2040|6100|360"
result.Add "PromotionStatus|6440|2040|2500|360"
result.Add "DateReported|9100|2040|2680|360"
result.Add "OccasionPeriodic|180|2970|2780|360"
result.Add "OccasionDetachment|3120|2970|2780|360"
result.Add "OccasionFrocking|6060|2970|2780|360"
result.Add "OccasionSpecial|9000|2970|2780|360"
result.Add "PeriodStart|180|3620|1900|360"
result.Add "PeriodEnd|2240|3620|1900|360"
result.Add "ObservationStatus|4300|3620|2400|360"
result.Add "ReportRegular|6860|3620|1500|360"
result.Add "ReportConcurrent|8520|3620|1600|360"
result.Add "ReportOperationalCommander|10280|3620|1500|360"
result.Add "ReadinessSummary|180|4550|3760|360"
result.Add "ReadinessDate|4100|4550|3760|360"
result.Add "BilletSubcategory|8020|4550|3760|360"
result.Add "ReportingSeniorName|180|5480|1800|360"
result.Add "ReportingSeniorGrade|2140|5480|1800|360"
result.Add "ReportingSeniorDesignator|4100|5480|1800|360"
result.Add "ReportingSeniorTitle|6060|5480|1800|360"
result.Add "ReportingSeniorUIC|8020|5480|1800|360"
result.Add "ReportingSeniorSSN|9980|5480|1800|360"
result.Add "CommandEmployment|180|6410|11600|950"
result.Add "PrimaryDuty|180|7930|5720|360"
result.Add "DutyDescription|6060|7930|5720|950"
result.Add "CounselingDate|180|9450|5720|360"
result.Add "CounselorName|6060|9450|5720|360"
result.Add "Trait33|180|10380|5000|360"
result.Add "Trait34|180|11000|5000|360"
result.Add "Trait35|180|11620|5000|360"
result.Add "Trait36|180|12240|5000|360"
result.Add "Trait37|180|12860|5000|360"
result.Add "Trait38|180|13480|5000|360"
result.Add "Trait39|180|14100|5000|360"
result.Add "@guidance|5550|10380|6000|4260"
Case "FITREP:1"
result.Add "AssignmentRecommendationA|180|180|5720|360"
result.Add "AssignmentRecommendationB|6060|180|5720|360"
result.Add "OfficialComments|180|1110|11600|3924"
result.Add "PromotionRecommendation|180|6584|3760|360"
result.Add "RaterName|4100|6584|3760|360"
result.Add "RaterDate|8020|6584|3760|360"
result.Add "ReportingSeniorAddress1|180|7514|3760|360"
result.Add "ReportingSeniorAddress2|4100|7514|3760|360"
result.Add "ReportingSeniorCity|8020|7514|3760|360"
result.Add "ReportingSeniorState|180|8444|3760|360"
result.Add "ReportingSeniorZip|4100|8444|3760|360"
result.Add "ReportingSeniorPhone|8020|8444|3760|360"
result.Add "ReportingSeniorDSN|180|9374|3760|360"
result.Add "StatementYes|4100|9374|3760|360"
result.Add "StatementNo|8020|9374|3760|360"
result.Add "RetentionYes|180|10304|3760|360"
result.Add "RetentionNo|4100|10304|3760|360"
result.Add "SeniorRaterName|8020|10304|3760|360"
result.Add "SeniorRaterDate|180|11234|11600|360"
End Select
Set ReportPageRows = result
End Function
