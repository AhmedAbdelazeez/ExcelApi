﻿// <auto-generated />
using ExcelApi.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

#nullable disable

namespace ExcelApi.Migrations
{
    [DbContext(typeof(Aplicationdbcontext))]
    partial class AplicationdbcontextModelSnapshot : ModelSnapshot
    {
        protected override void BuildModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .HasAnnotation("ProductVersion", "7.0.0")
                .HasAnnotation("Relational:MaxIdentifierLength", 128);

            SqlServerModelBuilderExtensions.UseIdentityColumns(modelBuilder);

            modelBuilder.Entity("ExcelApi.Models.testtable", b =>
                {
                    b.Property<string>("id")
                        .HasColumnType("nvarchar(450)");

                    b.Property<decimal>("area")
                        .HasColumnType("decimal(18,2)");

                    b.Property<string>("name")
                        .IsRequired()
                        .HasColumnType("nvarchar(max)");

                    b.HasKey("id");

                    b.ToTable("tests");
                });
#pragma warning restore 612, 618
        }
    }
}
