using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Tools.Migrations
{
    /// <inheritdoc />
    public partial class ChangeinExtraEnvelope : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "NRDataId",
                table: "ExtrasEnvelope");

            migrationBuilder.AddColumn<string>(
                name: "CatchNo",
                table: "ExtrasEnvelope",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "CatchNo",
                table: "ExtrasEnvelope");

            migrationBuilder.AddColumn<int>(
                name: "NRDataId",
                table: "ExtrasEnvelope",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }
    }
}
