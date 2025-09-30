using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Tools.Migrations
{
    /// <inheritdoc />
    public partial class Envelopemaking : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "BoxBreaking",
                table: "ProjectConfigs",
                newName: "EnvelopeMakingCriteria");

            migrationBuilder.AddColumn<string>(
                name: "BoxBreakingCriteria",
                table: "ProjectConfigs",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "BoxBreakingCriteria",
                table: "ProjectConfigs");

            migrationBuilder.RenameColumn(
                name: "EnvelopeMakingCriteria",
                table: "ProjectConfigs",
                newName: "BoxBreaking");
        }
    }
}
