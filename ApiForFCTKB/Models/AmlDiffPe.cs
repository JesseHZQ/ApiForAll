    

using System;

namespace ApiForFCTKB
{
    public partial class AmlDiffPe
    {
       public string TeradynePartNumber { get; set; }
       public string Manufacturer { get; set; }
       public string CorpStandardDescBefore { get; set; }
       public string CorpStandardDescAfter { get; set; }
       public bool CorpStandardDescChanged { get; set; }
       public string ParentPackageNameBefore { get; set; }
       public string ParentPackageNameAfter { get; set; }
       public bool ParentPackageNameChanged { get; set; }
       public string PrimAttTechnologyBefore { get; set; }
       public string PrimAttTechnologyAfter { get; set; }
       public bool PrimAttTechnologyChanged { get; set; }
       public string AqueousCompatibleBefore { get; set; }
       public string AqueousCompatibleAfter { get; set; }
       public bool AqueousCompatibleChanged { get; set; }
       public string SolventSensitiveBefore { get; set; }
       public string SolventSensitiveAfter { get; set; }
       public bool SolventSensitiveChanged { get; set; }
       public string MoistureSensitivityRatingBefore { get; set; }
       public string MoistureSensitivityRatingAfter { get; set; }
       public bool MoistureSensitivityRatingChanged { get; set; }
       public string IsRohsCompliantBefore { get; set; }
       public string IsRohsCompliantAfter { get; set; }
       public bool IsRohsCompliantChanged { get; set; }
       public string MaxProcTempInDegBefore { get; set; }
       public string MaxProcTempInDegAfter { get; set; }
       public bool MaxProcTempInDegChanged { get; set; }
       public string DurationMaxTempSecBefore { get; set; }
       public string DurationMaxTempSecAfter { get; set; }
       public bool DurationMaxTempSecChanged { get; set; }
       public string LeadFinishBefore { get; set; }
       public string LeadFinishAfter { get; set; }
       public bool LeadFinishChanged { get; set; }
       public string FoundryProcessCommentsBefore { get; set; }
       public string FoundryProcessCommentsAfter { get; set; }
       public bool FoundryProcessCommentsChanged { get; set; }
       public string T0ElecstatdiscdmBefore { get; set; }
       public string T0ElecstatdiscdmAfter { get; set; }
       public bool T0ElecstatdiscdmChanged { get; set; }
       public string T0ElecstatdishbmBefore { get; set; }
       public string T0ElecstatdishbmAfter { get; set; }
       public bool T0ElecstatdishbmChanged { get; set; }
       public string T0ElecstatdismmBefore { get; set; }
       public string T0ElecstatdismmAfter { get; set; }
       public bool T0ElecstatdismmChanged { get; set; }
       public int DiffFlag { get; set; }
       public bool IsNew { get; set; }
       public bool IsDelete { get; set; }
       public bool IsChange { get; set; }
    }

}