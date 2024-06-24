using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using NPOI.SS.Formula.Functions;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IndiaEventsWebApi.Models.MasterSheets
{
    public class HCPMaster
    {
        public string? RecordType { get; set; }
        public string? External_Id { get; set; }
        public string? AttendeeTypeCode { get; set; }
        public string? CurrencyCode { get; set; }
        public string? LastName { get; set; }
        public string? HCPName { get; set; }
        public string? Company_Name { get; set; }
        public string? Title { get; set; }
        public string? InActive { get; set; }
        public string? Employee_Group { get; set; }
        public string? Speciality { get; set; }
        public string? Medical_Council_Registration { get; set; }
        public string? GOorNGO { get; set; }
        public string? MISCode { get; set; }
        public string? AttendenceCountry { get; set; }
        public string? Hcp_Level { get; set; }
    }
    public class HCPMaster1
    {

        public string? LastName { get; set; }
        public string? FirstName { get; set; }
        public string? HCPName { get; set; }
        public string? Speciality { get; set; }
        public string? GOorNGO { get; set; }
        public string? MISCode { get; set; }
        public string? CompanyName { get; set; }

    }
    public class MIsandType
    {
        public string? MISCode { get; set; }
        public string? Type { get; set; }
    }
    public static class HcpCacheData
    {
        public static string HcpData = "HcpData";
    }
    public class HcpMasterData
    {
        [Column("LastName")]
        public string LastName { get; set; }
        [Column("FirstName")]
        public string FirstName { get; set; }
        [Column("HCPName")]
        public string HCPName { get; set; }
        [Column("HCP Type")]
        public string HCPType { get; set; }
        [Column("Employee Group")]
        public int? EmployeeGroup { get; set; }
        [Column("External ID")]
        public string ExternalID { get; set; }
        [Column("MIsCode")]
        public string MIsCode { get; set; }
        [Column("Company Name")]
        public string CompanyName { get; set; }
        [Column("Medical Council Registration")]
        public string MedicalCouncilRegistration { get; set; }
        [Column("Speciality")]
        public string Speciality { get; set; }
        [Column("FCPA Sign Off Date")]
        public string FCPASignOffDate { get; set; }
        [Column("FCPA Expiry Date")]
        public string FCPAExpiryDate { get; set; }
        [Column("FCPA Valid?")]
        public string FCPAValid { get; set; }
        [Column("ID")]
        [Key]
        public int? ID { get; set; }
    }
}
