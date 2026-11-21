"""
src/ui.py — Reusable UI Components styled with Tailwind CSS CDN
Provides professional hero headers, metric KPI cards, and insight banners
via Streamlit components iframe without requiring Node/npm.
"""
import streamlit as st
import streamlit.components.v1 as components

TAILWIND_CDN = '<link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">'

def render_hero_header(title: str, subtitle: str, metadata: str = "P071 | MCA DA ABES"):
    """Render an executive Hero Banner styled with Tailwind CSS."""
    html_code = f"""
    {TAILWIND_CDN}
    <div class="bg-gradient-to-r from-blue-900 via-indigo-900 to-blue-800 text-white p-6 rounded-xl shadow-lg border border-blue-700 font-sans">
        <div class="flex items-center justify-between">
            <div>
                <h1 class="text-2xl font-bold tracking-tight text-white mb-1">📊 {title}</h1>
                <p class="text-blue-200 text-sm">{subtitle}</p>
            </div>
            <div class="hidden sm:block text-right">
                <span class="inline-block bg-blue-700 bg-opacity-60 text-blue-100 text-xs px-3 py-1 rounded-full border border-blue-500 font-mono">
                    {metadata}
                </span>
            </div>
        </div>
    </div>
    """
    components.html(html_code, height=110)

def render_kpi_cards(total_jobs: int, avg_salary_inr: float, top_city: str, top_role: str):
    """Render 4 responsive executive KPI Metric Cards in a single Tailwind grid."""
    avg_sal_lakhs = avg_salary_inr / 100000.0
    html_code = f"""
    {TAILWIND_CDN}
    <div class="grid grid-cols-2 md:grid-cols-4 gap-4 font-sans">
        <!-- Card 1 -->
        <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-200 hover:shadow-md transition">
            <div class="text-xs font-semibold uppercase tracking-wider text-gray-500">Total Postings</div>
            <div class="text-2xl font-extrabold text-blue-900 mt-1">{total_jobs:,}</div>
            <div class="text-xs text-green-600 mt-1 font-medium">↑ Verified Dataset</div>
        </div>
        <!-- Card 2 -->
        <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-200 hover:shadow-md transition">
            <div class="text-xs font-semibold uppercase tracking-wider text-gray-500">Average Salary</div>
            <div class="text-2xl font-extrabold text-indigo-900 mt-1">₹{avg_sal_lakhs:.1f} LPA</div>
            <div class="text-xs text-blue-600 mt-1 font-medium">Benchmark Avg</div>
        </div>
        <!-- Card 3 -->
        <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-200 hover:shadow-md transition">
            <div class="text-xs font-semibold uppercase tracking-wider text-gray-500">Top Hiring City</div>
            <div class="text-2xl font-extrabold text-blue-950 mt-1 truncate">{top_city}</div>
            <div class="text-xs text-purple-600 mt-1 font-medium">Prime Hotspot</div>
        </div>
        <!-- Card 4 -->
        <div class="bg-white p-4 rounded-xl shadow-sm border border-gray-200 hover:shadow-md transition">
            <div class="text-xs font-semibold uppercase tracking-wider text-gray-500">Most Demanded Role</div>
            <div class="text-2xl font-extrabold text-blue-950 mt-1 truncate">{top_role}</div>
            <div class="text-xs text-amber-600 mt-1 font-medium">Highest Volume</div>
        </div>
    </div>
    """
    components.html(html_code, height=115)

def render_insight_box(insight_text: str, title: str = "Key Business Insight"):
    """Render a highlighted contextual insight callout banner."""
    html_code = f"""
    {TAILWIND_CDN}
    <div class="bg-blue-50 border-l-4 border-blue-600 p-4 rounded-r-lg shadow-xs font-sans">
        <div class="flex items-start">
            <span class="text-blue-600 font-bold mr-2 text-base">💡 {title}:</span>
            <p class="text-blue-900 text-sm leading-relaxed">{insight_text}</p>
        </div>
    </div>
    """
    components.html(html_code, height=75)
