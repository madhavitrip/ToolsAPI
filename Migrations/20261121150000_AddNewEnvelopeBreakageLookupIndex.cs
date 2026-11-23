using ERPToolsAPI.Data;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Migrations;

namespace Tools.Migrations
{
    [DbContext(typeof(ERPToolsDbContext))]
    [Migration("20261121150000_AddNewEnvelopeBreakageLookupIndex")]
    public partial class AddNewEnvelopeBreakageLookupIndex : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateIndex(
                name: "IX_newenvelopebreakage_ProjectId_CenterListId_NrDataId",
                table: "newenvelopebreakage",
                columns: new[] { "ProjectId", "CenterListId", "NrDataId" });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_newenvelopebreakage_ProjectId_CenterListId_NrDataId",
                table: "newenvelopebreakage");
        }
    }
}
