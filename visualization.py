# visualization.py

import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Any

class MarketShareVisualizer:
    """
    Handles all visualization components for market share analysis.

    Attributes:
        color_palette (List[Tuple[float, float, float]]): List of RGB tuples for consistent coloring.
    """

    def __init__(self, palette: str = "husl", num_colors: int = 8):
        """
        Initializes the MarketShareVisualizer with a specific color palette.

        Args:
            palette (str): Seaborn color palette name.
            num_colors (int): Number of distinct colors to generate.
        """
        # Set style for consistent look
        plt.style.use('seaborn-darkgrid')
        sns.set_palette(palette, n_colors=num_colors)
        self.color_palette = sns.color_palette(palette, n_colors=num_colors)

    def create_market_share_chart(
        self,
        data: Dict[str, float],
        title: str = "Market Share Distribution",
        explode_first: bool = True
    ) -> Optional[FigureCanvas]:
        """
        Create a pie chart for market share visualization.

        Args:
            data (Dict[str, float]): Dictionary with brand names as keys and their market shares as values.
            title (str): Title of the pie chart.
            explode_first (bool): Whether to explode the first slice for emphasis.

        Returns:
            Optional[FigureCanvas]: Matplotlib canvas containing the pie chart or None if data is invalid.
        """
        if not data:
            print("No data provided for market share chart.")
            return None

        # Sort data by value for better visualization
        sorted_data = dict(sorted(data.items(), key=lambda x: x[1], reverse=True))
        labels = list(sorted_data.keys())
        sizes = list(sorted_data.values())

        # Define explode
        explode = [0.05] + [0] * (len(labels) - 1) if explode_first else [0] * len(labels)

        # Define colors
        colors = self._assign_colors(labels)

        # Create figure
        fig, ax = plt.subplots(figsize=(8, 8))
        wedges, texts, autotexts = ax.pie(
            sizes,
            labels=labels,
            autopct='%1.1f%%',
            colors=colors,
            explode=explode,
            startangle=140,
            wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
        )

        # Equal aspect ratio ensures that pie is drawn as a circle.
        ax.axis('equal')
        ax.set_title(title, fontsize=14, pad=20)

        # Enhance text visibility
        plt.setp(autotexts, size=10, weight="bold")
        plt.setp(texts, size=10)

        # Create canvas
        canvas = FigureCanvas(fig)

        # Close the figure to free memory
        plt.close(fig)

        return canvas

    def create_regional_analysis_chart(
        self,
        df: pd.DataFrame,
        region_col: str,
        value_col: str,
        title: str = "Regional Distribution",
        kind: str = "bar"
    ) -> Optional[FigureCanvas]:
        """
        Create a bar or pie chart for regional analysis.

        Args:
            df (pd.DataFrame): DataFrame containing regional data.
            region_col (str): Column name for regions.
            value_col (str): Column name for values to plot.
            title (str): Title of the chart.
            kind (str): Type of chart ('bar', 'pie').

        Returns:
            Optional[FigureCanvas]: Matplotlib canvas containing the chart or None if data is invalid.
        """
        if df.empty:
            print("Empty DataFrame provided for regional analysis chart.")
            return None

        # Aggregate data
        aggregated = df.groupby(region_col)[value_col].sum().reset_index()

        if kind == "bar":
            fig, ax = plt.subplots(figsize=(12, 6))
            sns.barplot(data=aggregated, x=region_col, y=value_col, palette=self.color_palette, ax=ax)
            ax.set_title(title, fontsize=14, pad=20)
            ax.set_xlabel(region_col, fontsize=12)
            ax.set_ylabel("Allocated Yearly", fontsize=12)
            plt.xticks(rotation=45)

            # Add value labels on bars
            for index, row in aggregated.iterrows():
                ax.text(row.name, row[value_col], f'{row[value_col]:.1f}', color='black', ha="center")

        elif kind == "pie":
            fig, ax = plt.subplots(figsize=(8, 8))
            labels = aggregated[region_col]
            sizes = aggregated[value_col]
            colors = self._assign_colors(labels)
            explode = [0.05] + [0] * (len(labels) - 1)

            wedges, texts, autotexts = ax.pie(
                sizes,
                labels=labels,
                autopct='%1.1f%%',
                colors=colors,
                explode=explode,
                startangle=140,
                wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
            )
            ax.axis('equal')
            ax.set_title(title, fontsize=14, pad=20)
            plt.setp(autotexts, size=10, weight="bold")
            plt.setp(texts, size=10)

        else:
            print(f"Unsupported chart type: {kind}")
            return None

        # Create canvas
        canvas = FigureCanvas(fig)

        # Close the figure to free memory
        plt.close(fig)

        return canvas

    def create_trend_chart(
        self,
        df: pd.DataFrame,
        time_col: str,
        brands: List[str],
        share_col: str = "Share",
        title: str = "Market Share Trends"
    ) -> Optional[FigureCanvas]:
        """
        Create a line chart for trend analysis.

        Args:
            df (pd.DataFrame): DataFrame containing trend data.
            time_col (str): Column name for the time period.
            brands (List[str]): List of brand names to plot.
            share_col (str): Column name for market share values.
            title (str): Title of the trend chart.

        Returns:
            Optional[FigureCanvas]: Matplotlib canvas containing the trend chart or None if data is invalid.
        """
        if df.empty or not brands:
            print("Insufficient data provided for trend analysis chart.")
            return None

        fig, ax = plt.subplots(figsize=(12, 6))

        for idx, brand in enumerate(brands):
            brand_data = df[df['Brand'] == brand]
            if brand_data.empty:
                continue
            ax.plot(
                brand_data[time_col],
                brand_data[share_col],
                marker='o',
                linewidth=2,
                label=brand,
                color=self.color_palette[idx % len(self.color_palette)]
            )

        ax.set_title(title, fontsize=14, pad=20)
        ax.set_xlabel("Time Period", fontsize=12)
        ax.set_ylabel("Market Share (%)", fontsize=12)
        ax.legend(title="Brands", fontsize=10, title_fontsize=12)
        ax.grid(True, linestyle='--', alpha=0.7)
        plt.xticks(rotation=45)

        plt.tight_layout()
        canvas = FigureCanvas(fig)
        plt.close(fig)
        return canvas

    def create_class_distribution_chart(
        self,
        df: pd.DataFrame,
        class_col: str,
        brand_col: str,
        share_col: str = "Share",
        title: str = "Class Distribution Analysis"
    ) -> Optional[FigureCanvas]:
        """
        Create a stacked bar chart for class distribution.

        Args:
            df (pd.DataFrame): DataFrame containing class distribution data.
            class_col (str): Column name for classes.
            brand_col (str): Column name for brands.
            share_col (str): Column name for share values.
            title (str): Title of the class distribution chart.

        Returns:
            Optional[FigureCanvas]: Matplotlib canvas containing the class distribution chart or None if data is invalid.
        """
        if df.empty:
            print("Empty DataFrame provided for class distribution chart.")
            return None

        # Pivot the data
        pivot_df = df.pivot_table(
            index=class_col,
            columns=brand_col,
            values=share_col,
            aggfunc='sum'
        ).fillna(0)

        if pivot_df.empty:
            print("Pivot table is empty for class distribution chart.")
            return None

        fig, ax = plt.subplots(figsize=(12, 6))
        pivot_df.plot(
            kind='bar',
            stacked=True,
            ax=ax,
            color=self.color_palette[:len(pivot_df.columns)]
        )

        ax.set_title(title, fontsize=14, pad=20)
        ax.set_xlabel("Class", fontsize=12)
        ax.set_ylabel("Market Share (%)", fontsize=12)
        ax.legend(title="Brands", bbox_to_anchor=(1.05, 1), loc='upper left')
        plt.xticks(rotation=45)

        # Add percentage labels
        for container in ax.containers:
            ax.bar_label(container, fmt='%.1f%%', label_type='center')

        plt.tight_layout()
        canvas = FigureCanvas(fig)
        plt.close(fig)
        return canvas

    def create_summary_dashboard(
        self,
        market_share: Dict[str, float],
        regional_data: pd.DataFrame,
        trend_data: pd.DataFrame,
        class_data: pd.DataFrame,
        time_col: str,
        brands: List[str]
    ) -> Optional[FigureCanvas]:
        """
        Create a comprehensive dashboard with multiple charts.

        Args:
            market_share (Dict[str, float]): Market share data.
            regional_data (pd.DataFrame): DataFrame for regional analysis.
            trend_data (pd.DataFrame): DataFrame for trend analysis.
            class_data (pd.DataFrame): DataFrame for class distribution.
            time_col (str): Column name for the time period in trend data.
            brands (List[str]): List of brands in trend data.

        Returns:
            Optional[FigureCanvas]: Matplotlib canvas containing the dashboard or None if any data is invalid.
        """
        if not market_share or regional_data.empty or trend_data.empty or class_data.empty:
            print("Insufficient data provided for summary dashboard.")
            return None

        fig = plt.figure(constrained_layout=True, figsize=(20, 15))
        gs = fig.add_gridspec(2, 2)

        # Market Share Pie Chart
        ax1 = fig.add_subplot(gs[0, 0])
        self._create_pie_subplot(ax1, market_share, "Market Share Distribution")

        # Regional Bar Chart
        ax2 = fig.add_subplot(gs[0, 1])
        self._create_regional_subplot(ax2, regional_data, "Regional Distribution")

        # Trend Line Chart
        ax3 = fig.add_subplot(gs[1, 0])
        self._create_trend_subplot(ax3, trend_data, time_col, brands, "Market Share Trends")

        # Class Distribution Stacked Bar Chart
        ax4 = fig.add_subplot(gs[1, 1])
        self._create_class_subplot(ax4, class_data, "Class Distribution Analysis")

        plt.tight_layout()
        canvas = FigureCanvas(fig)
        plt.close(fig)
        return canvas

    def _create_pie_subplot(
        self,
        ax: Any,
        data: Dict[str, float],
        title: str = "Market Share Distribution"
    ):
        """
        Helper method to create a pie chart subplot.

        Args:
            ax (matplotlib.axes.Axes): Matplotlib Axes object to plot on.
            data (Dict[str, float]): Data for the pie chart.
            title (str): Title of the subplot.
        """
        if not data:
            ax.text(0.5, 0.5, 'No data available', horizontalalignment='center', verticalalignment='center')
            ax.axis('off')
            return

        # Sort data
        sorted_data = dict(sorted(data.items(), key=lambda x: x[1], reverse=True))
        labels = list(sorted_data.keys())
        sizes = list(sorted_data.values())
        explode = [0.05] + [0] * (len(labels) - 1)
        colors = self._assign_colors(labels)

        wedges, texts, autotexts = ax.pie(
            sizes,
            labels=labels,
            autopct='%1.1f%%',
            colors=colors,
            explode=explode,
            startangle=140,
            wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
        )
        ax.axis('equal')
        ax.set_title(title, fontsize=14, pad=20)
        plt.setp(autotexts, size=10, weight="bold")
        plt.setp(texts, size=10)

    def _create_regional_subplot(
        self,
        ax: Any,
        df: pd.DataFrame,
        title: str = "Regional Distribution"
    ):
        """
        Helper method to create a bar chart subplot for regional analysis.

        Args:
            ax (matplotlib.axes.Axes): Matplotlib Axes object to plot on.
            df (pd.DataFrame): DataFrame containing regional data.
            title (str): Title of the subplot.
        """
        if df.empty:
            ax.text(0.5, 0.5, 'No regional data available', horizontalalignment='center', verticalalignment='center')
            ax.axis('off')
            return

        sns.barplot(data=df, x='Region', y='Share', palette=self.color_palette, ax=ax)
        ax.set_title(title, fontsize=14, pad=20)
        ax.set_xlabel("Region", fontsize=12)
        ax.set_ylabel("Market Share (%)", fontsize=12)
        plt.xticks(rotation=45)

        # Add value labels on bars
        for index, row in df.iterrows():
            ax.text(index, row['Share'], f'{row["Share"]:.1f}%', color='black', ha="center")

    def _create_trend_subplot(
        self,
        ax: Any,
        df: pd.DataFrame,
        time_col: str,
        brands: List[str],
        title: str = "Market Share Trends"
    ):
        """
        Helper method to create a line chart subplot for trend analysis.

        Args:
            ax (matplotlib.axes.Axes): Matplotlib Axes object to plot on.
            df (pd.DataFrame): DataFrame containing trend data.
            time_col (str): Column name for time periods.
            brands (List[str]): List of brands to plot.
            title (str): Title of the subplot.
        """
        if df.empty or not brands:
            ax.text(0.5, 0.5, 'No trend data available', horizontalalignment='center', verticalalignment='center')
            ax.axis('off')
            return

        for idx, brand in enumerate(brands):
            brand_data = df[df['Brand'] == brand]
            if brand_data.empty:
                continue
            ax.plot(
                brand_data[time_col],
                brand_data['Share'],
                marker='o',
                linewidth=2,
                label=brand,
                color=self.color_palette[idx % len(self.color_palette)]
            )

        ax.set_title(title, fontsize=14, pad=20)
        ax.set_xlabel("Time Period", fontsize=12)
        ax.set_ylabel("Market Share (%)", fontsize=12)
        ax.legend(title="Brands", fontsize=10, title_fontsize=12)
        ax.grid(True, linestyle='--', alpha=0.7)
        plt.xticks(rotation=45)

    def _create_class_subplot(
        self,
        ax: Any,
        df: pd.DataFrame,
        title: str = "Class Distribution Analysis"
    ):
        """
        Helper method to create a stacked bar chart subplot for class distribution.

        Args:
            ax (matplotlib.axes.Axes): Matplotlib Axes object to plot on.
            df (pd.DataFrame): DataFrame containing class distribution data.
            title (str): Title of the subplot.
        """
        if df.empty:
            ax.text(0.5, 0.5, 'No class distribution data available', horizontalalignment='center', verticalalignment='center')
            ax.axis('off')
            return

        # Pivot the data
        pivot_df = df.pivot_table(
            index='Class',
            columns='Brand',
            values='Share',
            aggfunc='sum'
        ).fillna(0)

        if pivot_df.empty:
            ax.text(0.5, 0.5, 'No class distribution data available', horizontalalignment='center', verticalalignment='center')
            ax.axis('off')
            return

        pivot_df.plot(
            kind='bar',
            stacked=True,
            ax=ax,
            color=self.color_palette[:len(pivot_df.columns)]
        )

        ax.set_title(title, fontsize=14, pad=20)
        ax.set_xlabel("Class", fontsize=12)
        ax.set_ylabel("Market Share (%)", fontsize=12)
        ax.legend(title="Brands", bbox_to_anchor=(1.05, 1), loc='upper left')
        plt.xticks(rotation=45)

        # Add percentage labels
        for container in ax.containers:
            ax.bar_label(container, fmt='%.1f%%', label_type='center')

    def _assign_colors(self, categories: List[str]) -> List[Any]:
        """
        Assign colors to categories ensuring consistency.

        Args:
            categories (List[str]): List of category names.

        Returns:
            List[Any]: List of color tuples.
        """
        color_map = {}
        assigned_colors = []
        for idx, category in enumerate(categories):
            if category not in color_map:
                color_map[category] = self.color_palette[idx % len(self.color_palette)]
            assigned_colors.append(color_map[category])
        return assigned_colors

    def create_heatmap(
        self,
        data: pd.DataFrame,
        title: str = "Heatmap Analysis",
        xlabel: str = "X-axis",
        ylabel: str = "Y-axis"
    ) -> Optional[FigureCanvas]:
        """
        Create a heatmap for data visualization.

        Args:
            data (pd.DataFrame): Pivoted DataFrame for heatmap.
            title (str): Title of the heatmap.
            xlabel (str): Label for X-axis.
            ylabel (str): Label for Y-axis.

        Returns:
            Optional[FigureCanvas]: Matplotlib canvas containing the heatmap or None if data is invalid.
        """
        if data.empty:
            print("Empty DataFrame provided for heatmap.")
            return None

        fig, ax = plt.subplots(figsize=(12, 8))
        sns.heatmap(data, annot=True, fmt=".1f", cmap="YlGnBu", ax=ax, linewidths=.5)
        ax.set_title(title, fontsize=14, pad=20)
        ax.set_xlabel(xlabel, fontsize=12)
        ax.set_ylabel(ylabel, fontsize=12)

        plt.tight_layout()
        canvas = FigureCanvas(fig)
        plt.close(fig)
        return canvas
