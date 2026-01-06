"""
Finance Reports - Financial report generators
"""

import logging
from typing import Dict, Any, List, Optional
from datetime import datetime

logger = logging.getLogger("office_bot.finance.reports")

class FinanceReports:
    """
    Generate financial report structures
    """
    
    def __init__(self):
        logger.info("Finance Reports initialized")
    
    def generate_balance_sheet(
        self,
        company_name: str,
        as_of_date: str,
        assets: Dict[str, Any],
        liabilities: Dict[str, Any],
        equity: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Generate Balance Sheet structure
        """
        
        # Calculate totals
        total_current_assets = sum(assets.get("current", {}).values())
        total_non_current_assets = sum(assets.get("non_current", {}).values())
        total_assets = total_current_assets + total_non_current_assets
        
        total_current_liabilities = sum(liabilities.get("current", {}).values())
        total_non_current_liabilities = sum(liabilities.get("non_current", {}).values())
        total_liabilities = total_current_liabilities + total_non_current_liabilities
        
        total_equity = sum(equity.values())
        
        return {
            "company_name": company_name,
            "report_type": "Balance Sheet",
            "as_of_date": as_of_date,
            "data": {
                "assets": {
                    "current": assets.get("current", {}),
                    "non_current": assets.get("non_current", {}),
                    "total_current": total_current_assets,
                    "total_non_current": total_non_current_assets,
                    "total": total_assets
                },
                "liabilities": {
                    "current": liabilities.get("current", {}),
                    "non_current": liabilities.get("non_current", {}),
                    "total_current": total_current_liabilities,
                    "total_non_current": total_non_current_liabilities,
                    "total": total_liabilities
                },
                "equity": {
                    "items": equity,
                    "total": total_equity
                },
                "total_liabilities_equity": total_liabilities + total_equity,
                "is_balanced": abs(total_assets - (total_liabilities + total_equity)) < 0.01
            }
        }
    
    def generate_income_statement(
        self,
        company_name: str,
        period: str,
        revenue: float,
        cogs: float,
        operating_expenses: Dict[str, float],
        other_income: float = 0,
        other_expense: float = 0,
        tax_rate: float = 0.22
    ) -> Dict[str, Any]:
        """
        Generate Income Statement structure
        """
        
        gross_profit = revenue - cogs
        total_operating_expenses = sum(operating_expenses.values())
        operating_income = gross_profit - total_operating_expenses
        income_before_tax = operating_income + other_income - other_expense
        tax = income_before_tax * tax_rate if income_before_tax > 0 else 0
        net_income = income_before_tax - tax
        
        return {
            "company_name": company_name,
            "report_type": "Income Statement",
            "period": period,
            "data": {
                "revenue": revenue,
                "cogs": cogs,
                "gross_profit": gross_profit,
                "gross_margin": (gross_profit / revenue * 100) if revenue > 0 else 0,
                "operating_expenses": operating_expenses,
                "total_operating_expenses": total_operating_expenses,
                "operating_income": operating_income,
                "operating_margin": (operating_income / revenue * 100) if revenue > 0 else 0,
                "other_income": other_income,
                "other_expense": other_expense,
                "income_before_tax": income_before_tax,
                "tax": tax,
                "tax_rate": tax_rate,
                "net_income": net_income,
                "net_margin": (net_income / revenue * 100) if revenue > 0 else 0
            }
        }
    
    def generate_cash_flow_statement(
        self,
        company_name: str,
        period: str,
        operating_activities: Dict[str, float],
        investing_activities: Dict[str, float],
        financing_activities: Dict[str, float],
        beginning_cash: float
    ) -> Dict[str, Any]:
        """
        Generate Cash Flow Statement structure
        """
        
        net_operating = sum(operating_activities.values())
        net_investing = sum(investing_activities.values())
        net_financing = sum(financing_activities.values())
        
        net_change = net_operating + net_investing + net_financing
        ending_cash = beginning_cash + net_change
        
        return {
            "company_name": company_name,
            "report_type": "Cash Flow Statement",
            "period": period,
            "data": {
                "operating_activities": {
                    "items": operating_activities,
                    "net": net_operating
                },
                "investing_activities": {
                    "items": investing_activities,
                    "net": net_investing
                },
                "financing_activities": {
                    "items": financing_activities,
                    "net": net_financing
                },
                "net_change_in_cash": net_change,
                "beginning_cash": beginning_cash,
                "ending_cash": ending_cash
            }
        }
