"""
Statistics Module - Statistical analysis utilities
"""

import logging
from typing import List, Dict, Any, Optional, Tuple
import statistics
from collections import Counter

logger = logging.getLogger("office_bot.analyst.statistics")

class StatisticsAnalyzer:
    """
    Statistical analysis for data
    """
    
    def __init__(self):
        logger.info("Statistics Analyzer initialized")
    
    def basic_stats(self, data: List[float]) -> Dict[str, float]:
        """
        Calculate basic statistics
        """
        if not data:
            return {}
        
        # Filter out None values
        clean_data = [x for x in data if x is not None]
        
        if not clean_data:
            return {}
        
        return {
            "count": len(clean_data),
            "sum": sum(clean_data),
            "mean": statistics.mean(clean_data),
            "median": statistics.median(clean_data),
            "min": min(clean_data),
            "max": max(clean_data),
            "range": max(clean_data) - min(clean_data),
            "std_dev": statistics.stdev(clean_data) if len(clean_data) > 1 else 0,
            "variance": statistics.variance(clean_data) if len(clean_data) > 1 else 0
        }
    
    def percentiles(
        self, 
        data: List[float], 
        percentiles: List[float] = [25, 50, 75, 90, 95, 99]
    ) -> Dict[str, float]:
        """
        Calculate percentiles
        """
        if not data:
            return {}
        
        clean_data = sorted([x for x in data if x is not None])
        n = len(clean_data)
        
        result = {}
        for p in percentiles:
            idx = int(n * p / 100)
            result[f"p{int(p)}"] = clean_data[min(idx, n-1)]
        
        return result
    
    def frequency_distribution(
        self, 
        data: List[Any], 
        top_n: int = 10
    ) -> Dict[Any, int]:
        """
        Calculate frequency distribution
        """
        counter = Counter(data)
        return dict(counter.most_common(top_n))
    
    def detect_outliers(
        self, 
        data: List[float], 
        method: str = "iqr"
    ) -> Dict[str, Any]:
        """
        Detect outliers using IQR method
        """
        if not data or len(data) < 4:
            return {"outliers": [], "bounds": {}}
        
        clean_data = sorted([x for x in data if x is not None])
        
        q1_idx = len(clean_data) // 4
        q3_idx = 3 * len(clean_data) // 4
        
        q1 = clean_data[q1_idx]
        q3 = clean_data[q3_idx]
        iqr = q3 - q1
        
        lower_bound = q1 - 1.5 * iqr
        upper_bound = q3 + 1.5 * iqr
        
        outliers = [x for x in clean_data if x < lower_bound or x > upper_bound]
        
        return {
            "outliers": outliers,
            "outlier_count": len(outliers),
            "bounds": {
                "lower": lower_bound,
                "upper": upper_bound,
                "q1": q1,
                "q3": q3,
                "iqr": iqr
            }
        }
    
    def correlation(self, x: List[float], y: List[float]) -> float:
        """
        Calculate Pearson correlation coefficient
        """
        if len(x) != len(y) or len(x) < 2:
            return 0
        
        n = len(x)
        mean_x = sum(x) / n
        mean_y = sum(y) / n
        
        numerator = sum((xi - mean_x) * (yi - mean_y) for xi, yi in zip(x, y))
        
        sum_sq_x = sum((xi - mean_x) ** 2 for xi in x)
        sum_sq_y = sum((yi - mean_y) ** 2 for yi in y)
        
        denominator = (sum_sq_x * sum_sq_y) ** 0.5
        
        if denominator == 0:
            return 0
        
        return numerator / denominator
    
    def trend_analysis(
        self, 
        data: List[float],
        periods: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """
        Analyze trend in time series data
        """
        if not data or len(data) < 2:
            return {"trend": "insufficient_data"}
        
        # Simple linear trend
        n = len(data)
        x = list(range(n))
        
        mean_x = sum(x) / n
        mean_y = sum(data) / n
        
        numerator = sum((xi - mean_x) * (yi - mean_y) for xi, yi in zip(x, data))
        denominator = sum((xi - mean_x) ** 2 for xi in x)
        
        slope = numerator / denominator if denominator != 0 else 0
        
        # Determine trend direction
        if slope > 0.01:
            trend = "increasing"
        elif slope < -0.01:
            trend = "decreasing"
        else:
            trend = "stable"
        
        # Calculate growth
        first_val = data[0] if data[0] != 0 else 1
        last_val = data[-1]
        total_growth = ((last_val - first_val) / abs(first_val)) * 100
        
        return {
            "trend": trend,
            "slope": slope,
            "total_growth_pct": total_growth,
            "start_value": data[0],
            "end_value": data[-1],
            "average": sum(data) / len(data)
      }
