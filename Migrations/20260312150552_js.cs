using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Tools.Migrations
{
    /// <inheritdoc />
    public partial class js : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_BoxBreakingResults_ProjectId_NrDataId_BoxNo",
                table: "BoxBreakingResults");

            migrationBuilder.DropColumn(
                name: "CenterSortModified",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "NodalCodeRef",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "NodalSortModified",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "RouteRef",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "RouteSortModified",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "CatchNo",
                table: "BoxBreakingResults");

            migrationBuilder.DropColumn(
                name: "ExtraId",
                table: "BoxBreakingResults");

            migrationBuilder.DropColumn(
                name: "NrDataId",
                table: "BoxBreakingResults");

            migrationBuilder.AlterColumn<int>(
                name: "NrDataId",
                table: "EnvelopeBreakingResults",
                type: "int",
                nullable: false,
                defaultValue: 0,
                oldClrType: typeof(int),
                oldType: "int",
                oldNullable: true);

            migrationBuilder.AddColumn<int>(
                name: "EnvelopeBreakingResultId",
                table: "BoxBreakingResults",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.CreateIndex(
                name: "IX_BoxBreakingResults_ProjectId_BoxNo",
                table: "BoxBreakingResults",
                columns: new[] { "ProjectId", "BoxNo" });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_BoxBreakingResults_ProjectId_BoxNo",
                table: "BoxBreakingResults");

            migrationBuilder.DropColumn(
                name: "EnvelopeBreakingResultId",
                table: "BoxBreakingResults");

            migrationBuilder.AlterColumn<int>(
                name: "NrDataId",
                table: "EnvelopeBreakingResults",
                type: "int",
                nullable: true,
                oldClrType: typeof(int),
                oldType: "int");

            migrationBuilder.AddColumn<int>(
                name: "CenterSortModified",
                table: "EnvelopeBreakingResults",
                type: "int",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "NodalCodeRef",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<double>(
                name: "NodalSortModified",
                table: "EnvelopeBreakingResults",
                type: "double",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "RouteRef",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<int>(
                name: "RouteSortModified",
                table: "EnvelopeBreakingResults",
                type: "int",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "CatchNo",
                table: "BoxBreakingResults",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<int>(
                name: "ExtraId",
                table: "BoxBreakingResults",
                type: "int",
                nullable: true);

            migrationBuilder.AddColumn<int>(
                name: "NrDataId",
                table: "BoxBreakingResults",
                type: "int",
                nullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_BoxBreakingResults_ProjectId_NrDataId_BoxNo",
                table: "BoxBreakingResults",
                columns: new[] { "ProjectId", "NrDataId", "BoxNo" });
        }
    }
}
