# visualization.py
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
import pandas as pd
import numpy as np
from typing import Dict, List, Optional

class MarketShareVisualizer:
    """Handles all visualization components for market share analysis."""

    def __init__(self):
        # Set style for consistent look
        plt.style.use('seaborn')
        sns.set_palette("husl")
        self.colors = sns.color_palette("husl", 8)

    def create_market_share_chart(
        self,
        data: Dict[str, float],
        title: str = "Market Share Distribution"
    ) -> FigureCanvasQTAgg:
        """Create pie chart for market share visualization."""
        fig, ax = plt.subplots(figsize=(10, 8))

        # Sort data by value for better visualization
        sorted_data = dict(sorted(data.items(), key=lambda x: x[1], reverse=True))

        wedges, texts, autotexts = ax.pie(
            sorted_data.values(),
            labels=sorted_data.keys(),
            autopct='%1.1f%%',
            colors=self.colors,
            explode=[0.05 if i == 0 else 0 for i in range(len(sorted_data))]
        )

        # Style enhancements
        plt.setp(autotexts, size=9, weight="bold")
        plt.setp(texts, size=10)
        ax.set_title(title, pad=20, size=14)

        return FigureCanvasQTAgg(fig)

    def create_regional_analysis_chart(
        self,
        df: pd.DataFrame,
        region_col: str,
        value_col: str,
        title: str = "Regional Distribution"
    ) -> FigureCanvasQTAgg:
        """Create bar chart for regional analysis."""
        fig, ax = plt.subplots(figsize=(12, 6))

        # Create bar plot with seaborn
        sns.barplot(
            data=df,
            x=region_col,
            y=value_col,
            ax=ax,
            palette=self.colors
        )

        # Enhance styling
        ax.set_title(title, pad=20, size=14)
        ax.set_xlabel(region_col, size=12)
        ax.set_ylabel("Market Share (%)", size=12)
        plt.xticks(rotation=45)

        # Add value labels on bars
        for i, v in enumerate(df[value_col]):
            ax.text(i, v, f'{v:.1f}%', ha='center', va='bottom')

        plt.tight_layout()
        return FigureCanvasQTAgg(fig)

    def create_trend_chart(
        self,
        df: pd.DataFrame,
        time_col: str,
        brands: List[str],
        title: str = "Market Share Trends"
    ) -> FigureCanvasQTAgg:
        """Create line chart for trend analysis."""
        fig, ax = plt.subplots(figsize=(12, 6))

        for idx, brand in enumerate(brands):
            brand_data = df[df['Brand'] == brand]
            ax.plot(
                brand_data[time_col],
                brand_data['Share'],
                marker='o',
                linewidth=2,
                label=brand,
                color=self.colors[idx]
            )

        ax.set_title(title, pad=20, size=14)
        ax.set_xlabel("Time Period", size=12)
        ax.set_ylabel("Market Share (%)", size=12)
        ax.legend(title="Brands", title_fontsize=12)
        ax.grid(True, linestyle='--', alpha=0.7)

        plt.tight_layout()
        return FigureCanvasQTAgg(fig)

    def create_class_distribution_chart(
        self,
        df: pd.DataFrame,
        title: str = "Class Distribution Analysis"
    ) -> FigureCanvasQTAgg:
        """Create stacked bar chart for class distribution."""
        fig, ax = plt.subplots(figsize=(12, 6))

        df_pivot = df.pivot_table(
            index='Class',
            columns='Brand',
            values='Share',
            aggfunc='sum'
        ).fillna(0)

        df_pivot.plot(
            kind='bar',
            stacked=True,
            ax=ax,
            color=self.colors
        )

        ax.set_title(title, pad=20, size=14)
        ax.set_xlabel("Class", size=12)
        ax.set_ylabel("Market Share (%)", size=12)
        ax.legend(title="Brands", bbox_to_anchor=(1.05, 1))

        plt.tight_layout()
        return FigureCanvasQTAgg(fig)

    def create_summary_dashboard(
        self,
        market_share: Dict[str, float],
        regional_data: pd.DataFrame,
        trend_data: pd.DataFrame,
        class_data: pd.DataFrame
    ) -> FigureCanvasQTAgg:
        """Create a comprehensive dashboard with multiple charts."""
        fig = plt.figure(figsize=(20, 12))

        # Market Share Pie Chart
        ax1 = fig.add_subplot(221)
        self._create_pie_subplot(ax1, market_share, "Market Share Distribution")

        # Regional Bar Chart
        ax2 = fig.add_subplot(222)
        self._create_regional_subplot(ax2, regional_data)

        # Trend Line Chart
        ax3 = fig.add_subplot(223)
        self._create_trend_subplot(ax3, trend_data)

        # Class Distribution
        ax4 = fig.add_subplot(224)
        self._create_class_subplot(ax4, class_data)

        plt.tight_layout()
        return FigureCanvasQTAgg(fig)