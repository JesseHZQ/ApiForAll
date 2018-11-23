﻿#pragma warning disable 1591
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CheckSlot.DataContext
{
	using System.Data.Linq;
	using System.Data.Linq.Mapping;
	using System.Data;
	using System.Collections.Generic;
	using System.Reflection;
	using System.Linq;
	using System.Linq.Expressions;
	using System.ComponentModel;
	using System;
	
	
	[global::System.Data.Linq.Mapping.DatabaseAttribute(Name="Planning_L2")]
	public partial class SlotDemandDataContext : System.Data.Linq.DataContext
	{
		
		private static System.Data.Linq.Mapping.MappingSource mappingSource = new AttributeMappingSource();
		
    #region 可扩展性方法定义
    partial void OnCreated();
    #endregion
		
		public SlotDemandDataContext() : 
				base(global::CheckSlot.Properties.Settings.Default.Planning_L2ConnectionString, mappingSource)
		{
			OnCreated();
		}
		
		public SlotDemandDataContext(string connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public SlotDemandDataContext(System.Data.IDbConnection connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public SlotDemandDataContext(string connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public SlotDemandDataContext(System.Data.IDbConnection connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public System.Data.Linq.Table<SlotDemandRequestDto> SlotDemandRequestDto
		{
			get
			{
				return this.GetTable<SlotDemandRequestDto>();
			}
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.SlotDemandRequestDto")]
	public partial class SlotDemandRequestDto
	{
		
		private int _Id;
		
		private string _Planner;
		
		private System.Nullable<int> _DemandId;
		
		private string _ItemNo;
		
		private string _DemandType;
		
		private System.Nullable<decimal> _RequiredQty;
		
		private System.Nullable<System.DateTime> _RequiredDate;
		
		private System.Nullable<decimal> _BegOfReqQty;
		
		private System.Nullable<decimal> _FullKitQty;
		
		private System.Nullable<decimal> _EndOfReqQty;
		
		private System.Nullable<bool> _IsHardAllocation;
		
		private System.Nullable<decimal> _SortOrder;
		
		private string _Slot;
		
		private string _WK;
		
		private string _Customer;
		
		private string _ItemDescription;
		
		private string _RefDocNo;
		
		private string _RefDocLine;
		
		private System.Nullable<bool> _IsPurple;
		
		private string _TERCommitWeek;
		
		public SlotDemandRequestDto()
		{
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Id", AutoSync=AutoSync.Always, DbType="Int NOT NULL IDENTITY", IsDbGenerated=true)]
		public int Id
		{
			get
			{
				return this._Id;
			}
			set
			{
				if ((this._Id != value))
				{
					this._Id = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Planner", DbType="VarChar(20)")]
		public string Planner
		{
			get
			{
				return this._Planner;
			}
			set
			{
				if ((this._Planner != value))
				{
					this._Planner = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_DemandId", DbType="Int")]
		public System.Nullable<int> DemandId
		{
			get
			{
				return this._DemandId;
			}
			set
			{
				if ((this._DemandId != value))
				{
					this._DemandId = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ItemNo", DbType="VarChar(50)")]
		public string ItemNo
		{
			get
			{
				return this._ItemNo;
			}
			set
			{
				if ((this._ItemNo != value))
				{
					this._ItemNo = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_DemandType", DbType="VarChar(20)")]
		public string DemandType
		{
			get
			{
				return this._DemandType;
			}
			set
			{
				if ((this._DemandType != value))
				{
					this._DemandType = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_RequiredQty", DbType="Decimal(18,4)")]
		public System.Nullable<decimal> RequiredQty
		{
			get
			{
				return this._RequiredQty;
			}
			set
			{
				if ((this._RequiredQty != value))
				{
					this._RequiredQty = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_RequiredDate", DbType="DateTime")]
		public System.Nullable<System.DateTime> RequiredDate
		{
			get
			{
				return this._RequiredDate;
			}
			set
			{
				if ((this._RequiredDate != value))
				{
					this._RequiredDate = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_BegOfReqQty", DbType="Decimal(18,4)")]
		public System.Nullable<decimal> BegOfReqQty
		{
			get
			{
				return this._BegOfReqQty;
			}
			set
			{
				if ((this._BegOfReqQty != value))
				{
					this._BegOfReqQty = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_FullKitQty", DbType="Decimal(18,4)")]
		public System.Nullable<decimal> FullKitQty
		{
			get
			{
				return this._FullKitQty;
			}
			set
			{
				if ((this._FullKitQty != value))
				{
					this._FullKitQty = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_EndOfReqQty", DbType="Decimal(18,4)")]
		public System.Nullable<decimal> EndOfReqQty
		{
			get
			{
				return this._EndOfReqQty;
			}
			set
			{
				if ((this._EndOfReqQty != value))
				{
					this._EndOfReqQty = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_IsHardAllocation", DbType="Bit")]
		public System.Nullable<bool> IsHardAllocation
		{
			get
			{
				return this._IsHardAllocation;
			}
			set
			{
				if ((this._IsHardAllocation != value))
				{
					this._IsHardAllocation = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_SortOrder", DbType="Decimal(18,4)")]
		public System.Nullable<decimal> SortOrder
		{
			get
			{
				return this._SortOrder;
			}
			set
			{
				if ((this._SortOrder != value))
				{
					this._SortOrder = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Slot", DbType="VarChar(200)")]
		public string Slot
		{
			get
			{
				return this._Slot;
			}
			set
			{
				if ((this._Slot != value))
				{
					this._Slot = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_WK", DbType="VarChar(20)")]
		public string WK
		{
			get
			{
				return this._WK;
			}
			set
			{
				if ((this._WK != value))
				{
					this._WK = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Customer", DbType="VarChar(200)")]
		public string Customer
		{
			get
			{
				return this._Customer;
			}
			set
			{
				if ((this._Customer != value))
				{
					this._Customer = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ItemDescription", DbType="VarChar(200)")]
		public string ItemDescription
		{
			get
			{
				return this._ItemDescription;
			}
			set
			{
				if ((this._ItemDescription != value))
				{
					this._ItemDescription = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_RefDocNo", DbType="VarChar(20)")]
		public string RefDocNo
		{
			get
			{
				return this._RefDocNo;
			}
			set
			{
				if ((this._RefDocNo != value))
				{
					this._RefDocNo = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_RefDocLine", DbType="VarChar(20)")]
		public string RefDocLine
		{
			get
			{
				return this._RefDocLine;
			}
			set
			{
				if ((this._RefDocLine != value))
				{
					this._RefDocLine = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_IsPurple", DbType="Bit")]
		public System.Nullable<bool> IsPurple
		{
			get
			{
				return this._IsPurple;
			}
			set
			{
				if ((this._IsPurple != value))
				{
					this._IsPurple = value;
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_TERCommitWeek", DbType="VarChar(50)")]
		public string TERCommitWeek
		{
			get
			{
				return this._TERCommitWeek;
			}
			set
			{
				if ((this._TERCommitWeek != value))
				{
					this._TERCommitWeek = value;
				}
			}
		}
	}
}
#pragma warning restore 1591