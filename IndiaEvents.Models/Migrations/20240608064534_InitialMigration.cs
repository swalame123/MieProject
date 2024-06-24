using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace IndiaEvents.Models.Migrations
{
    /// <inheritdoc />
    public partial class InitialMigration : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "vendorCodeGenerations",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    InitiatorNameName = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    InitiatorEmail = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    VendorAccount = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    MisCode = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    BenificiaryName = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    PanCardName = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    PanNumber = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    BankAccountNumber = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    BankName = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    IfscCode = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    SwiftCode = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    IbnNumber = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Email = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    PanCardDocument = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    ChequeDocument = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    TaxResidenceCertificate = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    FinanceChecker = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_vendorCodeGenerations", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "vendorCodeGenerations");
        }
    }
}
