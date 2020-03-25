using Microsoft.EntityFrameworkCore.Migrations;

namespace WebApplication1.Migrations
{
    public partial class Indexduzenleme : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AlterColumn<string>(
                name: "FileName",
                table: "CellRecords",
                nullable: true,
                oldClrType: typeof(string),
                oldNullable: true);

            migrationBuilder.CreateIndex(
                name: "IX_CellRecords_RowIndex_ColumnIndex_FileName_TableIndex",
                table: "CellRecords",
                columns: new[] { "RowIndex", "ColumnIndex", "FileName", "TableIndex" },
                unique: true,
                filter: "[FileName] IS NOT NULL");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_CellRecords_RowIndex_ColumnIndex_FileName_TableIndex",
                table: "CellRecords");

            migrationBuilder.AlterColumn<string>(
                name: "FileName",
                table: "CellRecords",
                nullable: true,
                oldClrType: typeof(string),
                oldNullable: true);
        }
    }
}
