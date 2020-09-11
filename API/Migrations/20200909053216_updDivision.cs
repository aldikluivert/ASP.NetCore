using Microsoft.EntityFrameworkCore.Migrations;

namespace API.Migrations
{
    public partial class updDivision : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_TB_M_Division_TB_M_Department_DepartmentId",
                table: "TB_M_Division");

            migrationBuilder.AlterColumn<int>(
                name: "DepartmentId",
                table: "TB_M_Division",
                nullable: true,
                oldClrType: typeof(int));

            migrationBuilder.AddForeignKey(
                name: "FK_TB_M_Division_TB_M_Department_DepartmentId",
                table: "TB_M_Division",
                column: "DepartmentId",
                principalTable: "TB_M_Department",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_TB_M_Division_TB_M_Department_DepartmentId",
                table: "TB_M_Division");

            migrationBuilder.AlterColumn<int>(
                name: "DepartmentId",
                table: "TB_M_Division",
                nullable: false,
                oldClrType: typeof(int),
                oldNullable: true);

            migrationBuilder.AddForeignKey(
                name: "FK_TB_M_Division_TB_M_Department_DepartmentId",
                table: "TB_M_Division",
                column: "DepartmentId",
                principalTable: "TB_M_Department",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }
    }
}
