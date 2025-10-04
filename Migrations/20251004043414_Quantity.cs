using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Tools.Migrations
{
    /// <inheritdoc />
    public partial class Quantity : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<int>(
                name: "Quantity",
                table: "NRDatas",
                type: "int",
                nullable: false,
                defaultValue: 0,
                oldClrType: typeof(int),
                oldType: "int",
                oldNullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_EnvelopeBreakages_NrDataId",
                table: "EnvelopeBreakages",
                column: "NrDataId",
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_EnvelopeBreakages_NrDataId",
                table: "EnvelopeBreakages");

            migrationBuilder.AlterColumn<int>(
                name: "Quantity",
                table: "NRDatas",
                type: "int",
                nullable: true,
                oldClrType: typeof(int),
                oldType: "int");
        }
    }
}
