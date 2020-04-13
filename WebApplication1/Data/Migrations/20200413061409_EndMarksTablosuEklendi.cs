using System;
using Microsoft.EntityFrameworkCore.Migrations;

namespace WebApplication1.Migrations
{
    public partial class EndMarksTablosuEklendi : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "EndMarks",
                columns: table => new
                {
                    Id = table.Column<Guid>(nullable: false),
                    TemplateName = table.Column<string>(nullable: true),
                    SheetIndex = table.Column<int>(nullable: false),
                    RowIndex = table.Column<int>(nullable: false),
                    ColumnIndex = table.Column<int>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_EndMarks", x => x.Id);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "EndMarks");
        }
    }
}
