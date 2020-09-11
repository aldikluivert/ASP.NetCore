using Microsoft.EntityFrameworkCore.Migrations;

namespace API.Migrations
{
    public partial class updateRole : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_Tb_M_UserRole_TB_M_Role_RoleId",
                table: "Tb_M_UserRole");

            migrationBuilder.DropPrimaryKey(
                name: "PK_Tb_M_UserRole",
                table: "Tb_M_UserRole");

            migrationBuilder.AlterColumn<string>(
                name: "RoleId",
                table: "Tb_M_UserRole",
                nullable: true,
                oldClrType: typeof(string));

            migrationBuilder.AddPrimaryKey(
                name: "PK_Tb_M_UserRole",
                table: "Tb_M_UserRole",
                column: "UserId");

            migrationBuilder.AddForeignKey(
                name: "FK_Tb_M_UserRole_TB_M_Role_RoleId",
                table: "Tb_M_UserRole",
                column: "RoleId",
                principalTable: "TB_M_Role",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_Tb_M_UserRole_TB_M_Role_RoleId",
                table: "Tb_M_UserRole");

            migrationBuilder.DropPrimaryKey(
                name: "PK_Tb_M_UserRole",
                table: "Tb_M_UserRole");

            migrationBuilder.AlterColumn<string>(
                name: "RoleId",
                table: "Tb_M_UserRole",
                nullable: false,
                oldClrType: typeof(string),
                oldNullable: true);

            migrationBuilder.AddPrimaryKey(
                name: "PK_Tb_M_UserRole",
                table: "Tb_M_UserRole",
                columns: new[] { "UserId", "RoleId" });

            migrationBuilder.AddForeignKey(
                name: "FK_Tb_M_UserRole_TB_M_Role_RoleId",
                table: "Tb_M_UserRole",
                column: "RoleId",
                principalTable: "TB_M_Role",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }
    }
}
