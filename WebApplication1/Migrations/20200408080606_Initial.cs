using System;
using Microsoft.EntityFrameworkCore.Migrations;

namespace WebApplication1.Migrations
{
    public partial class Initial : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "CellRecords",
                columns: table => new
                {
                    Id = table.Column<Guid>(nullable: false),
                    RowIndex = table.Column<int>(nullable: false),
                    ColumnIndex = table.Column<int>(nullable: false),
                    Data = table.Column<string>(nullable: true),
                    TableIndex = table.Column<int>(nullable: false),
                    TemplateName = table.Column<string>(nullable: true),
                    FileName = table.Column<string>(nullable: true),
                    Date = table.Column<DateTime>(nullable: false),
                    Type = table.Column<string>(nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_CellRecords", x => x.Id);
                });

            migrationBuilder.CreateIndex(
                name: "IX_CellRecords_RowIndex_ColumnIndex_FileName_TableIndex",
                table: "CellRecords",
                columns: new[] { "RowIndex", "ColumnIndex", "FileName", "TableIndex" },
                unique: true,
                filter: "[FileName] IS NOT NULL");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "CellRecords");
        }
    }
}
