using System;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Tools.Migrations
{
    /// <inheritdoc />
    public partial class AddStatusToBreakingResults : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_EnvelopeBreakages_NrDataId",
                table: "EnvelopeBreakages");

            migrationBuilder.RenameColumn(
                name: "NRQuantity",
                table: "EnvelopeBreakingResults",
                newName: "DistrictSort");

            migrationBuilder.AddColumn<bool>(
                name: "Status",
                table: "Projects",
                type: "tinyint(1)",
                nullable: false,
                defaultValue: false);

            migrationBuilder.AddColumn<string>(
                name: "MssAttached",
                table: "ProjectConfigs",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "MssTypes",
                table: "ProjectConfigs",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<bool>(
                name: "RoundOffBeforeEnhancement",
                table: "ProjectConfigs",
                type: "tinyint(1)",
                nullable: false,
                defaultValue: false);

            migrationBuilder.AddColumn<string>(
                name: "Day",
                table: "NRDatas",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "District",
                table: "NRDatas",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<int>(
                name: "DistrictSort",
                table: "NRDatas",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "EnvLotNo",
                table: "NRDatas",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "LotNo",
                table: "NRDatas",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "NRDataId",
                table: "NRDatas",
                type: "int",
                nullable: true);

            migrationBuilder.AddColumn<bool>(
                name: "Status",
                table: "NRDatas",
                type: "tinyint(1)",
                nullable: false,
                defaultValue: false);

            migrationBuilder.AddColumn<int>(
                name: "Steps",
                table: "NRDatas",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<string>(
                name: "UploadList",
                table: "NRDatas",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "ParentModuleIds",
                table: "Modules",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "NodalCode",
                table: "ExtrasEnvelope",
                type: "varchar(255)",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<int>(
                name: "Status",
                table: "ExtrasEnvelope",
                type: "int",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<string>(
                name: "nodalValue",
                table: "ExtraConfigurations",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AlterColumn<string>(
                name: "EnvQuantity",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: false,
                oldClrType: typeof(int),
                oldType: "int")
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "District",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<string>(
                name: "OmrSerial",
                table: "EnvelopeBreakingResults",
                type: "longtext",
                nullable: true)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<bool>(
                name: "Status",
                table: "EnvelopeBreakingResults",
                type: "tinyint(1)",
                nullable: false,
                defaultValue: false);

            migrationBuilder.AddColumn<string>(
                name: "BookletSerial",
                table: "BoxBreakingResults",
                type: "longtext",
                nullable: false)
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.AddColumn<bool>(
                name: "Status",
                table: "BoxBreakingResults",
                type: "tinyint(1)",
                nullable: false,
                defaultValue: false);

            migrationBuilder.CreateTable(
                name: "ChangedNRData",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("MySql:ValueGenerationStrategy", MySqlValueGenerationStrategy.IdentityColumn),
                    ProjectId = table.Column<int>(type: "int", nullable: false),
                    CourseName = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    SubjectName = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    CenterCode = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    Quantity = table.Column<int>(type: "int", nullable: false),
                    NRQuantity = table.Column<int>(type: "int", nullable: false),
                    CatchNo = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    ExamDate = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    ExamTime = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    NRDatas = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    NodalCode = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    Pages = table.Column<int>(type: "int", nullable: false),
                    Route = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    RouteSort = table.Column<int>(type: "int", nullable: false),
                    CenterSort = table.Column<double>(type: "double", nullable: false),
                    NodalSort = table.Column<double>(type: "double", nullable: false),
                    Symbol = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4")
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ChangedNRData", x => x.Id);
                })
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.CreateTable(
                name: "EnvelopeLotReports",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("MySql:ValueGenerationStrategy", MySqlValueGenerationStrategy.IdentityColumn),
                    ProjectId = table.Column<int>(type: "int", nullable: false),
                    TemplateId = table.Column<int>(type: "int", nullable: false),
                    TemplateName = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    EnvLotNumbers = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    FileName = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    GeneratedAt = table.Column<DateTime>(type: "datetime(6)", nullable: false),
                    GeneratedBy = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    FilePath = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4")
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_EnvelopeLotReports", x => x.Id);
                    table.ForeignKey(
                        name: "FK_EnvelopeLotReports_Projects_ProjectId",
                        column: x => x.ProjectId,
                        principalTable: "Projects",
                        principalColumn: "ProjectId",
                        onDelete: ReferentialAction.Cascade);
                })
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.CreateTable(
                name: "MExtraConfigurations",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("MySql:ValueGenerationStrategy", MySqlValueGenerationStrategy.IdentityColumn),
                    GroupId = table.Column<int>(type: "int", nullable: false),
                    TypeId = table.Column<int>(type: "int", nullable: false),
                    ExtraType = table.Column<int>(type: "int", nullable: false),
                    Mode = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    Value = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    EnvelopeType = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    RangeConfig = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4")
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_MExtraConfigurations", x => x.Id);
                })
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.CreateTable(
                name: "MProjectConfigs",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("MySql:ValueGenerationStrategy", MySqlValueGenerationStrategy.IdentityColumn),
                    Modules = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    TypeId = table.Column<int>(type: "int", nullable: false),
                    GroupId = table.Column<int>(type: "int", nullable: false),
                    Envelope = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    BoxBreakingCriteria = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    DuplicateRemoveFields = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    SortingBoxReport = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    EnvelopeMakingCriteria = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    BoxCapacity = table.Column<int>(type: "int", nullable: false),
                    DuplicateCriteria = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    Enhancement = table.Column<double>(type: "double", nullable: false),
                    BoxNumber = table.Column<int>(type: "int", nullable: false),
                    OmrSerialNumber = table.Column<int>(type: "int", nullable: false),
                    ResetOnSymbolChange = table.Column<bool>(type: "tinyint(1)", nullable: false),
                    IsInnerBundlingDone = table.Column<bool>(type: "tinyint(1)", nullable: false),
                    InnerBundlingCriteria = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    ResetOmrSerialOnCatchChange = table.Column<bool>(type: "tinyint(1)", nullable: false),
                    BookletSerialNumber = table.Column<int>(type: "int", nullable: true),
                    ResetBookletSerialOnCatchChange = table.Column<bool>(type: "tinyint(1)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_MProjectConfigs", x => x.Id);
                })
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.CreateTable(
                name: "Mss",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("MySql:ValueGenerationStrategy", MySqlValueGenerationStrategy.IdentityColumn),
                    MssType = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    TypeId = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Mss", x => x.Id);
                })
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.CreateTable(
                name: "RPTMappings",
                columns: table => new
                {
                    ID = table.Column<int>(type: "int", nullable: false)
                        .Annotation("MySql:ValueGenerationStrategy", MySqlValueGenerationStrategy.IdentityColumn),
                    TemplateId = table.Column<int>(type: "int", nullable: false),
                    MappingJson = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4")
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_RPTMappings", x => x.ID);
                })
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.CreateTable(
                name: "RPTTemplates",
                columns: table => new
                {
                    TemplateId = table.Column<int>(type: "int", nullable: false)
                        .Annotation("MySql:ValueGenerationStrategy", MySqlValueGenerationStrategy.IdentityColumn),
                    GroupId = table.Column<int>(type: "int", nullable: true),
                    TypeId = table.Column<int>(type: "int", nullable: false),
                    ProjectId = table.Column<int>(type: "int", nullable: true),
                    UploadedByUserId = table.Column<int>(type: "int", nullable: true),
                    ModuleIds = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    TemplateName = table.Column<string>(type: "longtext", nullable: false)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    RPTFilePath = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    ParsedFieldsJson = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    DesignSnapshotJson = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    RequiredFieldsJson = table.Column<string>(type: "longtext", nullable: true)
                        .Annotation("MySql:CharSet", "utf8mb4"),
                    Version = table.Column<int>(type: "int", nullable: false),
                    CreatedDate = table.Column<DateTime>(type: "datetime(6)", nullable: false),
                    UpdatedDate = table.Column<DateTime>(type: "datetime(6)", nullable: false),
                    IsActive = table.Column<bool>(type: "tinyint(1)", nullable: true),
                    IsDeleted = table.Column<bool>(type: "tinyint(1)", nullable: true),
                    ReportStatus = table.Column<bool>(type: "tinyint(1)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_RPTTemplates", x => x.TemplateId);
                })
                .Annotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.CreateIndex(
                name: "IX_Projects_ProjectId",
                table: "Projects",
                column: "ProjectId",
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_ProjectConfigs_ProjectId",
                table: "ProjectConfigs",
                column: "ProjectId",
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_ExtrasEnvelope_NodalCode_ExtraId_ProjectId",
                table: "ExtrasEnvelope",
                columns: new[] { "NodalCode", "ExtraId", "ProjectId" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_ExtraConfigurations_ExtraType_ProjectId",
                table: "ExtraConfigurations",
                columns: new[] { "ExtraType", "ProjectId" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_EnvelopeBreakages_NrDataId_ProjectId",
                table: "EnvelopeBreakages",
                columns: new[] { "NrDataId", "ProjectId" },
                unique: true);

            migrationBuilder.CreateIndex(
                name: "IX_EnvelopeLotReports_ProjectId",
                table: "EnvelopeLotReports",
                column: "ProjectId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "ChangedNRData");

            migrationBuilder.DropTable(
                name: "EnvelopeLotReports");

            migrationBuilder.DropTable(
                name: "MExtraConfigurations");

            migrationBuilder.DropTable(
                name: "MProjectConfigs");

            migrationBuilder.DropTable(
                name: "Mss");

            migrationBuilder.DropTable(
                name: "RPTMappings");

            migrationBuilder.DropTable(
                name: "RPTTemplates");

            migrationBuilder.DropIndex(
                name: "IX_Projects_ProjectId",
                table: "Projects");

            migrationBuilder.DropIndex(
                name: "IX_ProjectConfigs_ProjectId",
                table: "ProjectConfigs");

            migrationBuilder.DropIndex(
                name: "IX_ExtrasEnvelope_NodalCode_ExtraId_ProjectId",
                table: "ExtrasEnvelope");

            migrationBuilder.DropIndex(
                name: "IX_ExtraConfigurations_ExtraType_ProjectId",
                table: "ExtraConfigurations");

            migrationBuilder.DropIndex(
                name: "IX_EnvelopeBreakages_NrDataId_ProjectId",
                table: "EnvelopeBreakages");

            migrationBuilder.DropColumn(
                name: "Status",
                table: "Projects");

            migrationBuilder.DropColumn(
                name: "MssAttached",
                table: "ProjectConfigs");

            migrationBuilder.DropColumn(
                name: "MssTypes",
                table: "ProjectConfigs");

            migrationBuilder.DropColumn(
                name: "RoundOffBeforeEnhancement",
                table: "ProjectConfigs");

            migrationBuilder.DropColumn(
                name: "Day",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "District",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "DistrictSort",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "EnvLotNo",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "LotNo",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "NRDataId",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "Status",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "Steps",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "UploadList",
                table: "NRDatas");

            migrationBuilder.DropColumn(
                name: "ParentModuleIds",
                table: "Modules");

            migrationBuilder.DropColumn(
                name: "NodalCode",
                table: "ExtrasEnvelope");

            migrationBuilder.DropColumn(
                name: "Status",
                table: "ExtrasEnvelope");

            migrationBuilder.DropColumn(
                name: "nodalValue",
                table: "ExtraConfigurations");

            migrationBuilder.DropColumn(
                name: "District",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "OmrSerial",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "Status",
                table: "EnvelopeBreakingResults");

            migrationBuilder.DropColumn(
                name: "BookletSerial",
                table: "BoxBreakingResults");

            migrationBuilder.DropColumn(
                name: "Status",
                table: "BoxBreakingResults");

            migrationBuilder.RenameColumn(
                name: "DistrictSort",
                table: "EnvelopeBreakingResults",
                newName: "NRQuantity");

            migrationBuilder.AlterColumn<int>(
                name: "EnvQuantity",
                table: "EnvelopeBreakingResults",
                type: "int",
                nullable: false,
                oldClrType: typeof(string),
                oldType: "longtext")
                .OldAnnotation("MySql:CharSet", "utf8mb4");

            migrationBuilder.CreateIndex(
                name: "IX_EnvelopeBreakages_NrDataId",
                table: "EnvelopeBreakages",
                column: "NrDataId",
                unique: true);
        }
    }
}
