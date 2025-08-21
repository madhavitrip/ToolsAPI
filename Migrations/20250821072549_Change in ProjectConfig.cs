using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Tools.Migrations
{
    /// <inheritdoc />
    public partial class ChangeinProjectConfig : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "PackagingSort",
                table: "ProjectConfigs",
                newName: "Extras");

            migrationBuilder.RenameColumn(
                name: "MergeKey",
                table: "ProjectConfigs",
                newName: "Envelope");

            migrationBuilder.RenameColumn(
                name: "EnvelopeSort",
                table: "ProjectConfigs",
                newName: "BoxBreaking");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "Extras",
                table: "ProjectConfigs",
                newName: "PackagingSort");

            migrationBuilder.RenameColumn(
                name: "Envelope",
                table: "ProjectConfigs",
                newName: "MergeKey");

            migrationBuilder.RenameColumn(
                name: "BoxBreaking",
                table: "ProjectConfigs",
                newName: "EnvelopeSort");
        }
    }
}
