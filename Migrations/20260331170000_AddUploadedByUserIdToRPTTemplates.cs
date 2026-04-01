using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Tools.Migrations
{
    /// <inheritdoc />
    public partial class AddUploadedByUserIdToRPTTemplates : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "UploadedByUserId",
                table: "RPTTemplates",
                type: "int",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "UploadedByUserId",
                table: "RPTTemplates");
        }
    }
}
