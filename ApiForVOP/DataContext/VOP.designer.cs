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

namespace ApiForVOP.DataContext
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
	
	
	[global::System.Data.Linq.Mapping.DatabaseAttribute(Name="DebugAssy")]
	public partial class VOPDataContext : System.Data.Linq.DataContext
	{
		
		private static System.Data.Linq.Mapping.MappingSource mappingSource = new AttributeMappingSource();
		
    #region 可扩展性方法定义
    partial void OnCreated();
    partial void InsertVOP_Result(VOP_Result instance);
    partial void UpdateVOP_Result(VOP_Result instance);
    partial void DeleteVOP_Result(VOP_Result instance);
    #endregion
		
		public VOPDataContext() : 
				base(global::ApiForVOP.Properties.Settings.Default.DebugAssyConnectionString, mappingSource)
		{
			OnCreated();
		}
		
		public VOPDataContext(string connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public VOPDataContext(System.Data.IDbConnection connection) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public VOPDataContext(string connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public VOPDataContext(System.Data.IDbConnection connection, System.Data.Linq.Mapping.MappingSource mappingSource) : 
				base(connection, mappingSource)
		{
			OnCreated();
		}
		
		public System.Data.Linq.Table<VOP_Result> VOP_Result
		{
			get
			{
				return this.GetTable<VOP_Result>();
			}
		}
	}
	
	[global::System.Data.Linq.Mapping.TableAttribute(Name="dbo.VOP_Result")]
	public partial class VOP_Result : INotifyPropertyChanging, INotifyPropertyChanged
	{
		
		private static PropertyChangingEventArgs emptyChangingEventArgs = new PropertyChangingEventArgs(String.Empty);
		
		private int _ID;
		
		private string _Date;
		
		private string _Detail;
		
		private string _PODetail;
		
		private System.Nullable<double> _TotalCost;
		
    #region 可扩展性方法定义
    partial void OnLoaded();
    partial void OnValidate(System.Data.Linq.ChangeAction action);
    partial void OnCreated();
    partial void OnIDChanging(int value);
    partial void OnIDChanged();
    partial void OnDateChanging(string value);
    partial void OnDateChanged();
    partial void OnDetailChanging(string value);
    partial void OnDetailChanged();
    partial void OnPODetailChanging(string value);
    partial void OnPODetailChanged();
    partial void OnTotalCostChanging(System.Nullable<double> value);
    partial void OnTotalCostChanged();
    #endregion
		
		public VOP_Result()
		{
			OnCreated();
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_ID", AutoSync=AutoSync.OnInsert, DbType="Int NOT NULL IDENTITY", IsPrimaryKey=true, IsDbGenerated=true)]
		public int ID
		{
			get
			{
				return this._ID;
			}
			set
			{
				if ((this._ID != value))
				{
					this.OnIDChanging(value);
					this.SendPropertyChanging();
					this._ID = value;
					this.SendPropertyChanged("ID");
					this.OnIDChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Date", DbType="VarChar(50)")]
		public string Date
		{
			get
			{
				return this._Date;
			}
			set
			{
				if ((this._Date != value))
				{
					this.OnDateChanging(value);
					this.SendPropertyChanging();
					this._Date = value;
					this.SendPropertyChanged("Date");
					this.OnDateChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_Detail", DbType="VarChar(MAX)")]
		public string Detail
		{
			get
			{
				return this._Detail;
			}
			set
			{
				if ((this._Detail != value))
				{
					this.OnDetailChanging(value);
					this.SendPropertyChanging();
					this._Detail = value;
					this.SendPropertyChanged("Detail");
					this.OnDetailChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_PODetail", DbType="VarChar(MAX)")]
		public string PODetail
		{
			get
			{
				return this._PODetail;
			}
			set
			{
				if ((this._PODetail != value))
				{
					this.OnPODetailChanging(value);
					this.SendPropertyChanging();
					this._PODetail = value;
					this.SendPropertyChanged("PODetail");
					this.OnPODetailChanged();
				}
			}
		}
		
		[global::System.Data.Linq.Mapping.ColumnAttribute(Storage="_TotalCost", DbType="Float")]
		public System.Nullable<double> TotalCost
		{
			get
			{
				return this._TotalCost;
			}
			set
			{
				if ((this._TotalCost != value))
				{
					this.OnTotalCostChanging(value);
					this.SendPropertyChanging();
					this._TotalCost = value;
					this.SendPropertyChanged("TotalCost");
					this.OnTotalCostChanged();
				}
			}
		}
		
		public event PropertyChangingEventHandler PropertyChanging;
		
		public event PropertyChangedEventHandler PropertyChanged;
		
		protected virtual void SendPropertyChanging()
		{
			if ((this.PropertyChanging != null))
			{
				this.PropertyChanging(this, emptyChangingEventArgs);
			}
		}
		
		protected virtual void SendPropertyChanged(String propertyName)
		{
			if ((this.PropertyChanged != null))
			{
				this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
			}
		}
	}
}
#pragma warning restore 1591
