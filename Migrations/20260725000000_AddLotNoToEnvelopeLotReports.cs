using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace ToolsAPI.Migrations
{
    /// <inheritdoc />
    public partial class AddLotNoToEnvelopeLotReports : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "LotNo",
                table: "EnvelopeLotReports",
                type: "int",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "LotNo",
                table: "EnvelopeLotReports");
        }
    }
}
