using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Tools.Migrations
{
    /// <inheritdoc />
    public partial class change : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<string>(
                name: "Env",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: true,
                oldClrType: typeof(string),
                oldType: "longtext")
                .Annotation("MySql:CharSet", "utf8mb4")
                .OldAnnotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AlterColumn<string>(
                name: "CatchNo",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: true,
                oldClrType: typeof(string),
                oldType: "longtext")
                .Annotation("MySql:CharSet", "utf8mb4")
                .OldAnnotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<int>(
                name: "CenterSort",
                table: "EnvelopeBreakingResults",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<string>(
                name: "CourseName",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "ExamDate",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "ExamTime",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<int>(
                name: "NRQuantity",
                table: "EnvelopeBreakingResults",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<string>(
                name: "NodalCode",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<double>(
                name: "NodalSort",
                table: "EnvelopeBreakingResults",
                type: "double",
                nullable: false,
                defaultValue: 0.0);

            migrationBuilder.AddColumn<int>(
                name: "Quantity",
                table: "EnvelopeBreakingResults",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "RouteSort",
                table: "EnvelopeBreakingResults",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "CenterSort",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "CourseName",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "ExamDate",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "ExamTime",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "NRQuantity",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "NodalCode",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "NodalSort",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "Quantity",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "RouteSort",
                table: "EnvelopeBreakingResults");

            migrationBuilder.UpdateData(
                table: "EnvelopeBreakingResults",
                keyColumn: "Env",
                keyValue: null,
                column: "Env",
                value: "");

            migrationBuilder.AlterColumn<string>(
                name: "Env",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: false,
                oldClrType: typeof(string),
                oldType: "longtext",
                oldNullable: true)
                .Annotation("MySql:CharSet", "utf8mb4")
                .OldAnnotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.UpdateData(
                table: "EnvelopeBreakingResults",
                keyColumn: "CatchNo",
                keyValue: null,
                column: "CatchNo",
                value: "");

            migrationBuilder.AlterColumn<string>(
                name: "CatchNo",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: false,
                oldClrType: typeof(string),
                oldType: "longtext",
                oldNullable: true)
                .Annotation("MySql:CharSet", "utf8mb4")
                .OldAnnotation("MySql:CharSet", "utf8mb4");
        }
    }
}
