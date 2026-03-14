using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Tools.Migrations
{
    /// <inheritdoc />
    public partial class changesconflictfields : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "CatchNo",
                table: "ConflictingFields");

            migrationBuilder.AddColumn<int>(
                name: "BookletSerialNumber",
                table: "ProjectConfigs",
                type: "int",
                nullable: true);

            migrationBuilder.AddColumn<bool>(
                name: "ResetBookletSerialOnCatchChange",
                table: "ProjectConfigs",
                type: "tinyint(1)",
                nullable: true);

            migrationBuilder.AlterColumn<double>(
                name: "NodalSort",
                table: "NRDatas",
                type: "double",
                nullable: false,
                oldClrType: typeof(int),
                oldType: "int");

            migrationBuilder.AlterColumn<double>(
                name: "CenterSort",
                table: "NRDatas",
                type: "double",
                nullable: false,
                oldClrType: typeof(int),
                oldType: "int");

            migrationBuilder.AddColumn<string>(
                name: "RangeConfig",
                table: "ExtraConfigurations",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AlterColumn<double>(
                name: "CenterSort",
                table: "EnvelopeBreakingResults",
                type: "double",
                nullable: false,
                oldClrType: typeof(int),
                oldType: "int");

            migrationBuilder.AlterColumn<int>(
                name: "Status",
                table: "ConflictingFields",
                type: "int",
                nullable: false,
                oldClrType: typeof(string),
                oldType: "longtext")
                .OldAnnotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<int>(
                name: "NRDataId",
                table: "ConflictingFields",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "BookletSerialNumber",
                table: "ProjectConfigs");

            migrationBuilder.DropColumn(
                name: "ResetBookletSerialOnCatchChange",
                table: "ProjectConfigs");

            migrationBuilder.DropColumn(
                name: "RangeConfig",
                table: "ExtraConfigurations");

            migrationBuilder.DropColumn(
                name: "NRDataId",
                table: "ConflictingFields");

            migrationBuilder.AlterColumn<int>(
                name: "NodalSort",
                table: "NRDatas",
                type: "int",
                nullable: false,
                oldClrType: typeof(double),
                oldType: "double");

            migrationBuilder.AlterColumn<int>(
                name: "CenterSort",
                table: "NRDatas",
                type: "int",
                nullable: false,
                oldClrType: typeof(double),
                oldType: "double");

            migrationBuilder.AlterColumn<int>(
                name: "CenterSort",
                table: "EnvelopeBreakingResults",
                type: "int",
                nullable: false,
                oldClrType: typeof(double),
                oldType: "double");

            migrationBuilder.AlterColumn<string>(
                name: "Status",
                table: "ConflictingFields",
                type: "longtext",
                nullable: false,
                oldClrType: typeof(int),
                oldType: "int")
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "CatchNo",
                table: "ConflictingFields",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");
        }
    }
}
