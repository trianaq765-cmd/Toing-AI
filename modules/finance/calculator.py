"""
Finance Calculator - Financial calculation utilities
"""

import logging
from typing import Dict, Any, List, Optional
from dataclasses import dataclass

logger = logging.getLogger("office_bot.finance.calculator")

@dataclass
class BEPResult:
    """Break Even Point result"""
    bep_units: float
    bep_rupiah: float
    contribution_margin: float
    fixed_cost: float
    variable_cost: float
    price: float

@dataclass
class DepreciationSchedule:
    """Depreciation schedule entry"""
    year: int
    depreciation: float
    accumulated: float
    book_value: float

class FinanceCalculator:
    """
    Financial calculations for business analysis
    """
    
    def __init__(self):
        logger.info("Finance Calculator initialized")
    
    # =========================================================================
    # BREAK EVEN POINT
    # =========================================================================
    
    def calculate_bep(
        self,
        fixed_cost: float,
        price_per_unit: float,
        variable_cost_per_unit: float
    ) -> BEPResult:
        """
        Calculate Break Even Point
        
        BEP (units) = Fixed Cost / (Price - Variable Cost)
        BEP (Rp) = BEP (units) x Price
        """
        contribution_margin = price_per_unit - variable_cost_per_unit
        
        if contribution_margin <= 0:
            raise ValueError("Price must be greater than variable cost")
        
        bep_units = fixed_cost / contribution_margin
        bep_rupiah = bep_units * price_per_unit
        
        return BEPResult(
            bep_units=bep_units,
            bep_rupiah=bep_rupiah,
            contribution_margin=contribution_margin,
            fixed_cost=fixed_cost,
            variable_cost=variable_cost_per_unit,
            price=price_per_unit
        )
    
    # =========================================================================
    # ROI & PROFITABILITY
    # =========================================================================
    
    def calculate_roi(self, investment: float, return_value: float) -> Dict[str, float]:
        """
        Calculate Return on Investment
        
        ROI = (Return - Investment) / Investment x 100%
        """
        profit = return_value - investment
        roi_pct = (profit / investment * 100) if investment > 0 else 0
        
        return {
            "investment": investment,
            "return": return_value,
            "profit": profit,
            "roi_percentage": roi_pct
        }
    
    def calculate_profit_margin(
        self,
        revenue: float,
        cost: float
    ) -> Dict[str, float]:
        """
        Calculate profit margins
        """
        gross_profit = revenue - cost
        gross_margin = (gross_profit / revenue * 100) if revenue > 0 else 0
        
        return {
            "revenue": revenue,
            "cost": cost,
            "gross_profit": gross_profit,
            "gross_margin_pct": gross_margin
        }
    
    # =========================================================================
    # DEPRECIATION
    # =========================================================================
    
    def calculate_straight_line_depreciation(
        self,
        cost: float,
        salvage_value: float,
        useful_life: int
    ) -> List[DepreciationSchedule]:
        """
        Calculate Straight Line Depreciation
        
        Annual Depreciation = (Cost - Salvage Value) / Useful Life
        """
        annual_depreciation = (cost - salvage_value) / useful_life
        
        schedule = []
        accumulated = 0
        
        for year in range(1, useful_life + 1):
            accumulated += annual_depreciation
            book_value = cost - accumulated
            
            schedule.append(DepreciationSchedule(
                year=year,
                depreciation=annual_depreciation,
                accumulated=accumulated,
                book_value=book_value
            ))
        
        return schedule
    
    def calculate_declining_balance_depreciation(
        self,
        cost: float,
        salvage_value: float,
        useful_life: int,
        rate: Optional[float] = None
    ) -> List[DepreciationSchedule]:
        """
        Calculate Declining Balance Depreciation
        
        Rate = 1 / Useful Life x 2 (double declining)
        Depreciation = Book Value x Rate
        """
        if rate is None:
            rate = (1 / useful_life) * 2  # Double declining
        
        schedule = []
        book_value = cost
        accumulated = 0
        
        for year in range(1, useful_life + 1):
            depreciation = book_value * rate
            
            # Don't depreciate below salvage value
            if book_value - depreciation < salvage_value:
                depreciation = book_value - salvage_value
            
            accumulated += depreciation
            book_value = cost - accumulated
            
            schedule.append(DepreciationSchedule(
                year=year,
                depreciation=depreciation,
                accumulated=accumulated,
                book_value=book_value
            ))
        
        return schedule
    
    # =========================================================================
    # FINANCIAL RATIOS
    # =========================================================================
    
    def calculate_liquidity_ratios(
        self,
        current_assets: float,
        current_liabilities: float,
        inventory: float = 0,
        cash: float = 0
    ) -> Dict[str, float]:
        """
        Calculate liquidity ratios
        """
        current_ratio = current_assets / current_liabilities if current_liabilities > 0 else 0
        quick_ratio = (current_assets - inventory) / current_liabilities if current_liabilities > 0 else 0
        cash_ratio = cash / current_liabilities if current_liabilities > 0 else 0
        
        return {
            "current_ratio": current_ratio,
            "quick_ratio": quick_ratio,
            "cash_ratio": cash_ratio
        }
    
    def calculate_solvency_ratios(
        self,
        total_debt: float,
        total_equity: float,
        total_assets: float
    ) -> Dict[str, float]:
        """
        Calculate solvency ratios
        """
        debt_to_equity = total_debt / total_equity if total_equity > 0 else 0
        debt_ratio = total_debt / total_assets if total_assets > 0 else 0
        
        return {
            "debt_to_equity": debt_to_equity,
            "debt_ratio": debt_ratio
        }
    
    def calculate_profitability_ratios(
        self,
        net_income: float,
        revenue: float,
        total_assets: float,
        total_equity: float
    ) -> Dict[str, float]:
        """
        Calculate profitability ratios
        """
        net_profit_margin = (net_income / revenue * 100) if revenue > 0 else 0
        roa = (net_income / total_assets * 100) if total_assets > 0 else 0
        roe = (net_income / total_equity * 100) if total_equity > 0 else 0
        
        return {
            "net_profit_margin": net_profit_margin,
            "return_on_assets": roa,
            "return_on_equity": roe
        }
