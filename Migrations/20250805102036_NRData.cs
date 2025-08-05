using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Tools.Migrations
{
    /// <inheritdoc />
    public partial class NRData : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "NR",
                table: "NRDatas",
                newName: "SubjectName");

            migrationBuilder.AddColumn<string>(
                name: "CatchNo",
                table: "NRDatas",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<int>(
                name: "CenterCode",
                table: "NRDatas",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<string>(
                name: "CourseName",
                table: "NRDatas",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "ExamDate",
                table: "NRDatas",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "ExamTime",
                table: "NRDatas",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "NRDatas",
                table: "NRDatas",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<int>(
                name: "Quantity",
                table: "NRDatas",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "CatchNo",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "CenterCode",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "CourseName",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "ExamDate",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "ExamTime",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "NRDatas",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "Quantity",
                table: "NRDatas");

            migrationBuilder.RenameColumn(
                name: "SubjectName",
                table: "NRDatas",
                newName: "NR");
        }
    }
}
