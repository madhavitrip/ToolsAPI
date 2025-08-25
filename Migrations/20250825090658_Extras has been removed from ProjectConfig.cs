using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Tools.Migrations
{
    /// <inheritdoc />
    public partial class ExtrashasbeenremovedfromProjectConfig : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Extras",
                table: "ProjectConfigs");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "Extras",
                table: "ProjectConfigs",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");
        }
    }
}
