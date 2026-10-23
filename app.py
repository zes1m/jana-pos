import os
import threading
from datetime import datetime, timedelta, timezone
import pandas as pd
from flask import Flask, request, jsonify, send_file, Response

app = Flask(__name__)

EXCEL_JANA = 'مبيعات_جنى_الوديان_الموحد.xlsx'
EXCEL_NATURE = 'مبيعات_جمال_الطبيعة.xlsx'
file_lock = threading.Lock()

with file_lock:
    if not os.path.exists(EXCEL_JANA):
        pd.DataFrame(columns=[
            'رقم الفاتورة', 'التاريخ والوقت', 'اسم البائع / الكاشير',
            'تفاصيل المواد المباعة', 'عدد القطع', 'المجموع الفرعي (د.ع)',
            'الخصم (د.ع)', 'الصافي المستلم (د.ع)', 'ملاحظات'
        ]).to_excel(EXCEL_JANA, index=False)
        
    if not os.path.exists(EXCEL_NATURE):
        pd.DataFrame(columns=[
            'رقم الفاتورة', 'التاريخ والوقت', 'اسم البائع / الكاشير',
            'تفاصيل المواد المباعة', 'عدد القطع', 'المجموع الفرعي (د.ع)',
            'الخصم (د.ع)', 'الصافي المستلم (د.ع)', 'ملاحظات'
        ]).to_excel(EXCEL_NATURE, index=False)

HTML_PAGE = """<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
  <title>نظام الكاشير الآمن - جنى الوديان & جمال الطبيعة</title>
  <style>
    :root { 
      --primary-jana: #14361b; --leaf-jana: #84b53a; 
      --primary-nature: #4a154b; --leaf-nature: #e01a4f; 
      --bg: #f4f7f4; 
    }
    * { box-sizing: border-box; font-family: system-ui, -apple-system, 'Segoe UI', Tahoma, sans-serif; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
    body { background: var(--bg); color: #222; padding-bottom: 95px; transition: 0.3s; }
    
    #lock-screen { position: fixed; inset: 0; background: rgba(15, 23, 42, 0.95); z-index: 9999; display: flex; align-items: center; justify-content: center; padding: 20px; }
    .lock-card { background: #fff; width: 100%; max-width: 380px; padding: 24px; border-radius: 16px; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.3); }
    .lock-card h2 { color: #1e293b; margin-bottom: 8px; font-size: 1.3rem; }
    .lock-card p { color: #64748b; font-size: 0.85rem; margin-bottom: 20px; }
    .lock-card input { width: 100%; padding: 12px; font-size: 1.1rem; text-align: center; border: 2px solid #cbd5e1; border-radius: 10px; margin-bottom: 15px; font-weight: bold; letter-spacing: 3px; }
    .lock-card button { width: 100%; padding: 12px; background: #0f172a; color: #fff; border: none; border-radius: 10px; font-weight: bold; font-size: 1rem; cursor: pointer; }
    .lock-error { color: #dc2626; font-size: 0.85rem; margin-top: 10px; display: none; font-weight: bold; }

    .header { background: #ffffff; padding: 12px 16px; border-bottom: 3px solid var(--current-leaf, #84b53a); position: sticky; top: 0; z-index: 25; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 10px rgba(0,0,0,0.06); }
    .brand-box { display: flex; align-items: center; gap: 12px; }
    .brand-title h1 { font-size: 1.15rem; color: var(--current-leaf, #84b53a); font-weight: 800; line-height: 1.2; }
    .brand-title span { font-size: 0.82rem; color: #555; font-weight: bold; }
    .actions-top { display: flex; gap: 6px; align-items: center; }
    .btn-switch { background: var(--current-primary, #14361b); color: #fff; border: none; padding: 8px 12px; border-radius: 8px; font-weight: bold; font-size: 0.78rem; cursor: pointer; display: flex; align-items: center; gap: 4px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    .btn-dl { background: #334155; color: #fff; text-decoration: none; padding: 8px 10px; border-radius: 8px; font-weight: bold; font-size: 0.78rem; display: none; align-items: center; gap: 4px; }
    
    .container { padding: 12px; max-width: 920px; margin: auto; }
    .top-bar { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 10px; justify-content: space-between; align-items: center; }
    .seller-box { flex: 1; min-width: 220px; background: #fff; padding: 8px 12px; border-radius: 10px; display: flex; align-items: center; gap: 8px; border: 1.5px solid #cbd5e1; }
    .seller-box label { font-weight: bold; font-size: 0.85rem; color: #333; white-space: nowrap; }
    .seller-box input { flex: 1; padding: 6px 8px; border: 1px solid #e2e8f0; border-radius: 6px; font-weight: bold; font-size: 0.88rem; }
    .admin-badge { background: #fef3c7; color: #d97706; padding: 6px 12px; border-radius: 8px; font-size: 0.8rem; font-weight: bold; border: 1px solid #fde68a; display: flex; align-items: center; gap: 6px; cursor: pointer; }
    
    .stats-bar { background: #e8f5e9; color: #14361b; padding: 9px 14px; border-radius: 10px; margin-bottom: 12px; font-size: 0.84rem; font-weight: bold; display: none; justify-content: space-between; border: 1px solid #a5d6a7; }
    
    .search-section { background: #fff; padding: 12px; border-radius: 12px; margin-bottom: 12px; border: 2px solid var(--current-leaf, #84b53a); box-shadow: 0 3px 8px rgba(0,0,0,0.05); }
    .search-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; font-size: 0.85rem; font-weight: bold; }
    .search-input-wrap { position: relative; display: flex; align-items: center; }
    .search-input-wrap input { width: 100%; padding: 11px 14px 11px 70px; border-radius: 9px; border: 1.5px solid #cbd5e1; font-size: 0.95rem; background: #f8fafc; font-weight: 600; }
    .btn-clear-search { position: absolute; left: 6px; background: #fee2e2; color: #dc2626; border: none; padding: 6px 10px; border-radius: 6px; font-size: 0.78rem; font-weight: bold; cursor: pointer; display: none; }
    
    .cats { display: flex; gap: 6px; overflow-x: auto; padding-bottom: 8px; margin-bottom: 10px; scrollbar-width: none; }
    .cats::-webkit-scrollbar { display: none; }
    .cat-btn { padding: 8px 14px; border: 1px solid #cbd5e1; background: #fff; border-radius: 20px; font-weight: bold; font-size: 0.8rem; white-space: nowrap; cursor: pointer; color: #333; transition: 0.15s; }
    .cat-btn.active { background: var(--current-primary, #14361b); color: #fff; border-color: var(--current-primary, #14361b); }
    
    .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(142px, 1fr)); gap: 9px; margin-bottom: 20px; }
    .card { background: #fff; border-radius: 11px; padding: 10px; border: 1.5px solid #e2e8f0; display: flex; flex-direction: column; justify-content: space-between; box-shadow: 0 2px 4px rgba(0,0,0,0.03); position: relative; }
    .card-badge { position: absolute; top: 6px; left: 6px; background: var(--current-leaf, #84b53a); color: #fff; font-weight: bold; font-size: 0.78rem; width: 24px; height: 24px; border-radius: 50%; display: none; align-items: center; justify-content: center; box-shadow: 0 2px 4px rgba(0,0,0,0.15); }
    .card h4 { font-size: 0.84rem; color: #222; margin-bottom: 6px; line-height: 1.35; padding-left: 20px; }
    .price { font-size: 0.92rem; font-weight: 800; color: #b08900; margin-bottom: 8px; }
    .btn-add { width: 100%; padding: 8px 4px; background: var(--current-primary, #14361b); color: #fff; border: none; border-radius: 7px; font-size: 0.8rem; font-weight: bold; cursor: pointer; }
    
    .cart { background: #fff; border-radius: 12px; padding: 15px; border: 2px solid var(--current-primary, #14361b); box-shadow: 0 4px 12px rgba(0,0,0,0.06); margin-bottom: 20px; }
    .item { display: flex; justify-content: space-between; align-items: center; padding: 8px 0; border-bottom: 1px solid #f1f3f5; font-size: 0.88rem; }
    .q-btn { width: 28px; height: 28px; border: none; background: #e2e8f0; border-radius: 6px; font-weight: bold; font-size: 1rem; cursor: pointer; }
    .row { display: flex; justify-content: space-between; align-items: center; margin-top: 8px; font-size: 0.92rem; }
    .total { font-size: 1.2rem; font-weight: bold; color: var(--current-primary, #14361b); border-top: 2px dashed #dee2e6; padding-top: 10px; margin-top: 10px; }
    .input { width: 100%; padding: 9px; border: 1px solid #cbd5e1; border-radius: 8px; margin-top: 10px; font-size: 0.88rem; }
    
    .percent-discounts { display: none; gap: 4px; margin-top: 6px; flex-wrap: wrap; }
    .percent-btn { flex: 1; padding: 5px 2px; background: #f1f5f9; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 0.75rem; font-weight: bold; cursor: pointer; color: #334155; text-align: center; }
    .percent-btn:hover { background: #e2e8f0; }

    .actions { display: flex; gap: 8px; margin-top: 12px; }
    .btn-save { flex: 2; padding: 13px; background: var(--current-primary, #14361b); color: #fff; border: none; border-radius: 8px; font-size: 0.95rem; font-weight: bold; cursor: pointer; }
    .btn-print { flex: 1; padding: 13px; background: #0284c7; color: #fff; border: none; border-radius: 8px; font-size: 0.95rem; font-weight: bold; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 4px; }
    .btn-clear { flex: 1; padding: 13px; background: #fee2e2; color: #dc2626; border: none; border-radius: 8px; font-weight: bold; cursor: pointer; }
    #toast { display: none; margin-top: 10px; background: #d1fae5; color: #065f46; padding: 10px; border-radius: 8px; text-align: center; font-weight: bold; border: 1px solid #34d399; }
    
    .dev-footer { background: #1e293b; color: #fff; text-align: center; padding: 14px 12px; border-radius: 12px; margin-top: 15px; font-size: 0.85rem; line-height: 1.6; }
    .dev-footer strong { color: #38bdf8; }
    
    .floating-bar { position: fixed; bottom: 0; left: 0; right: 0; background: var(--current-primary, #14361b); color: #fff; padding: 10px 16px; display: flex; justify-content: space-between; align-items: center; border-top: 2px solid var(--current-leaf, #84b53a); z-index: 30; box-shadow: 0 -2px 10px rgba(0,0,0,0.18); }
    .floating-bar button { background: var(--current-leaf, #84b53a); color: #fff; border: none; padding: 9px 16px; border-radius: 8px; font-weight: bold; font-size: 0.88rem; cursor: pointer; }

    /* تنسيق خاص للطباعة */
    @media print {
      body * { visibility: hidden; }
      #print-area, #print-area * { visibility: visible; }
      #print-area { position: absolute; left: 0; top: 0; width: 100%; padding: 20px; background: #fff; color: #000; direction: rtl; }
    }
  </style>
</head>
<body>

  <div id="lock-screen">
    <div class="lock-card">
      <h2>🔐 تسجيل دخول النظام</h2>
      <p>أدخل رمز المرور الخاص للإدارة أو اضغط دخول كـ (كاشير)</p>
      <input type="password" id="pin-input" placeholder="الرمز السري..." maxlength="6">
      <button onclick="verifyPin()">دخول الإدارة (أدمن)</button>
      <button onclick="enterAsCashier()" style="background:#64748b; margin-top:8px;">دخول سريع (كاشير عادي)</button>
      <div class="lock-error" id="lock-error">❌ الرمز السري غير صحيح!</div>
    </div>
  </div>

  <div class="header">
    <div class="brand-box">
      <div class="brand-title">
        <h1 id="app-title">جنى الوديان اللبنانية</h1>
        <span id="app-subtitle">نظام الكاشير الذكي</span>
      </div>
    </div>
    <div class="actions-top">
      <button class="btn-switch" onclick="switchCompany()">🔄 التبديل للثانية</button>
      <a href="#" id="dl-link" class="btn-dl">📥 إكسل المبيعات</a>
    </div>
  </div>

  <div class="container">
    <div class="top-bar">
      <div class="seller-box">
        <label>👤 اسم البائع:</label>
        <input type="text" id="seller-name" placeholder="اكتب اسمك..." oninput="localStorage.setItem('active_seller', this.value)">
      </div>
      <div class="admin-badge" id="role-badge" onclick="lockSystem()">🔒 وضع: كاشير</div>
    </div>

    <div class="stats-bar" id="stats-box">
      <span id="stat-total">إجمالي المبيعات: 0 د.ع</span>
      <span id="stat-count">عدد الفواتير: 0</span>
    </div>

    <div class="search-section">
      <div class="search-header">
        <span id="search-title">🔍 البحث السريع في المنتجات</span>
        <span id="search-results-count" style="color:var(--current-leaf);">يعرض المنتجات</span>
      </div>
      <div class="search-input-wrap">
        <input type="text" id="search-input" placeholder="ابحث باسم المنتج أو السعر..." oninput="onSearchInput()">
        <button class="btn-clear-search" id="btn-clear-search" onclick="clearSearch()">مسح ✖</button>
      </div>
    </div>

    <div class="cats" id="cats-container"></div>
    <div class="grid" id="grid"></div>

    <div class="cart" id="cart-section">
      <div style="font-weight:bold; border-bottom:2px solid #f1f3f5; padding-bottom:6px; margin-bottom:8px; display:flex; justify-content:space-between;">
        <span>🧾 تفاصيل الفاتورة الحالية</span>
        <span id="count" style="font-size:0.85rem;">0 قطع</span>
      </div>
      <div id="list"><p style="color:#888; text-align:center; margin:14px 0;">اضغط على المنتجات بالأعلى لإضافتها للفاتورة</p></div>

      <div class="row"><span>المجموع الفرعي:</span><span id="sub">0 د.ع</span></div>
      <div class="row" style="flex-direction:column; align-items:stretch;">
        <div style="display:flex; justify-content:space-between; align-items:center; width:100%;">
          <span>خصم مباشر (د.ع):</span>
          <input type="number" id="disc" value="0" step="500" style="width:130px; padding:6px; text-align:center; border:1px solid #ccc; border-radius:6px; font-weight:bold;" oninput="renderCart()">
        </div>
        <div class="percent-discounts" id="percent-discounts-box">
          <button class="percent-btn" onclick="applyPercentDiscount(10)">خصم 10%</button>
          <button class="percent-btn" onclick="applyPercentDiscount(20)">خصم 20%</button>
          <button class="percent-btn" onclick="applyPercentDiscount(30)">خصم 30%</button>
          <button class="percent-btn" onclick="applyPercentDiscount(40)">خصم 40%</button>
          <button class="percent-btn" onclick="applyPercentDiscount(50)">خصم 50%</button>
        </div>
      </div>
      <div class="row total"><span>الصافي للدفع:</span><span id="tot">0 د.ع</span></div>

      <input type="text" id="notes" class="input" placeholder="ملاحظات إضافية أو اسم الزبون (اختياري)...">

      <div class="actions">
        <button class="btn-save" id="btn-save" onclick="submitOrder()">💾 حفظ الفاتورة</button>
        <button class="btn-print" onclick="printReceipt()">🖨️ طباعة</button>
        <button class="btn-clear" onclick="clearCart()">تصفير</button>
      </div>
      <div id="toast"></div>
    </div>

    <div class="dev-footer">
      <div>تصميم وتطوير المبرمج <strong>مصطفى عبد الرزاق</strong></div>
      <a href="tel:07759093905" style="color:#38bdf8; text-decoration:none;">📞 07759093905</a>
    </div>
  </div>

  <!-- منطقة الفاتورة الخفية المخصصة للطباعة -->
  <div id="print-area" style="display:none;">
    <div style="text-align:center; margin-bottom:15px;">
      <h2 id="print-comp-name" style="font-size:1.4rem; margin-bottom:4px;"></h2>
      <p style="font-size:0.85rem; color:#555;">إيصال مبيعات رسمي</p>
      <hr style="border:0; border-top:1px dashed #000; margin:8px 0;">
    </div>
    <div style="font-size:0.9rem; margin-bottom:10px;">
      <div><strong>التاريخ:</strong> <span id="print-date"></span></div>
      <div><strong>البائع/الكاشير:</strong> <span id="print-seller"></span></div>
    </div>
    <table style="width:100%; border-collapse:collapse; font-size:0.88rem; margin-bottom:10px;">
      <thead>
        <tr style="border-bottom:1px solid #000;">
          <th style="text-align:right; padding:4px;">المادة</th>
          <th style="text-align:center; padding:4px;">الكمية</th>
          <th style="text-align:left; padding:4px;">المجموع</th>
        </tr>
      </thead>
      <tbody id="print-items-body"></tbody>
    </table>
    <hr style="border:0; border-top:1px dashed #000; margin:8px 0;">
    <div style="font-size:0.9rem;">
      <div style="display:flex; justify-content:space-between;"><span>المجموع الفرعي:</span> <span id="print-sub"></span></div>
      <div style="display:flex; justify-content:space-between;"><span>الخصم:</span> <span id="print-disc"></span></div>
      <div style="display:flex; justify-content:space-between; font-weight:bold; font-size:1.1rem; margin-top:5px;"><span>الصافي المطلوب:</span> <span id="print-tot"></span></div>
    </div>
    <div style="text-align:center; margin-top:20px; font-size:0.8rem; color:#555;">
      <p>شكراً لزيارتكم!</p>
      <p>تصميم: مصطفى عبد الرزاق</p>
    </div>
  </div>

  <div class="floating-bar">
    <div>
      <span style="font-size:0.78rem; opacity:0.9;">الصافي الحالي:</span><br>
      <strong id="float-tot" style="font-size:1.1rem; color:#fff;">0 د.ع</strong>
      <span id="float-qty" style="font-size:0.8rem; margin-right:6px;">(0 قطع)</span>
    </div>
    <button onclick="document.getElementById('cart-section').scrollIntoView({behavior:'smooth'})">إتمام وحفظ الفاتورة ⬇️</button>
  </div>

  <script>
    var ADMIN_PIN = "2026"; 
    var isAdmin = false;

    document.getElementById('seller-name').value = localStorage.getItem('active_seller') || '';
    var currentCompany = localStorage.getItem('jw_active_company') || 'jana';

    var productsJana = [
      { id: 1, name: "زيت الزيتون 250 مل", price: 4000, cat: "oils" },
      { id: 2, name: "زيت الزيتون 500 مل", price: 8000, cat: "oils" },
      { id: 3, name: "زيت الزيتون 3 لتر", price: 35000, cat: "oils" },
      { id: 4, name: "زيت الزيتون 10 لتر", price: 100000, cat: "oils" },
      { id: 5, name: "ماء الزهر العضوي", price: 3000, cat: "distillates" },
      { id: 6, name: "ماء الورد العضوي", price: 3000, cat: "distillates" },
      { id: 7, name: "ماء الشمار العضوي", price: 3000, cat: "distillates" },
      { id: 8, name: "ماء القصعين / المرمية", price: 3000, cat: "distillates" },
      { id: 9, name: "ماء الزعتر العضوي", price: 3000, cat: "distillates" },
      { id: 10, name: "ماء المركدوش العضوي", price: 3000, cat: "distillates" },
      { id: 11, name: "شراب الورد المركز", price: 4000, cat: "drinks" },
      { id: 12, name: "شراب التوت المركز", price: 4000, cat: "drinks" },
      { id: 13, name: "شراب الجلاب", price: 4000, cat: "drinks" },
      { id: 14, name: "شراب تمر هندي", price: 4000, cat: "drinks" },
      { id: 15, name: "شراب الرمان", price: 4000, cat: "drinks" },
      { id: 16, name: "شراب الفراولة", price: 4000, cat: "drinks" },
      { id: 17, name: "شراب النعنع", price: 4000, cat: "drinks" },
      { id: 18, name: "مربى التين", price: 3000, cat: "jams" },
      { id: 19, name: "مربى المشمش", price: 3000, cat: "jams" },
      { id: 20, name: "مربى التوت", price: 3000, cat: "jams" },
      { id: 21, name: "مربى الكرز", price: 3000, cat: "jams" },
      { id: 22, name: "مربى الفراولة", price: 3000, cat: "jams" },
      { id: 23, name: "خل البلسمك", price: 2000, cat: "vinegar" },
      { id: 24, name: "خل العنب", price: 2500, cat: "vinegar" },
      { id: 25, name: "خل تفاح", price: 2500, cat: "vinegar" },
      { id: 26, name: "دبس رمان الأمير 250 مل", price: 2000, cat: "vinegar" },
      { id: 27, name: "دبس رمان 500 مل بلاستك", price: 3000, cat: "vinegar" },
      { id: 28, name: "دبس رمان زجاجي 500 مل", price: 4000, cat: "vinegar" },
      { id: 29, name: "سماق 400 غرام", price: 6000, cat: "spices" },
      { id: 30, name: "زعتر نابلسي 250 غرام", price: 6000, cat: "spices" },
      { id: 31, name: "زعتر بالصنوبر 250 غرام", price: 6000, cat: "spices" },
      { id: 32, name: "زعتر اكسترا لبناني 400 غرام", price: 8000, cat: "spices" },
      { id: 33, name: "شطة حلوة 600 غرام", price: 6000, cat: "spices" },
      { id: 34, name: "شطة حارة 600 غرام", price: 5000, cat: "spices" },
      { id: 35, name: "مكدوس باذنجان بالجوز محشي لبنة 600غ", price: 10000, cat: "pickles" },
      { id: 36, name: "مكدوس باذنجان بالجوز 600 غرام", price: 8000, cat: "pickles" },
      { id: 37, name: "كبيس كورنيشون 600 غرام", price: 5000, cat: "pickles" },
      { id: 38, name: "كبيس لفت 600 غرام", price: 3000, cat: "pickles" },
      { id: 39, name: "كبيس مشكل", price: 3000, cat: "pickles" },
      { id: 40, name: "كبيس خيار", price: 3000, cat: "pickles" },
      { id: 41, name: "زيتون مشوي 600 غرام", price: 8000, cat: "olives" },
      { id: 42, name: "زيتون أخضر محشي لبنة 600 غرام", price: 8000, cat: "olives" },
      { id: 43, name: "زيتون أخضر محشي خضار 600 غرام", price: 8000, cat: "olives" },
      { id: 44, name: "زيتون أخضر بالزعتر الأخضر 600 غرام", price: 6000, cat: "olives" },
      { id: 45, name: "زيتون أخضر محشي جزر 600 غرام", price: 6000, cat: "olives" },
      { id: 46, name: "زيتون أسود شرائح 600 غرام", price: 5000, cat: "olives" },
      { id: 47, name: "زيتون أخضر شرائح", price: 5000, cat: "olives" },
      { id: 48, name: "زيتون أخضر حب 600 غرام", price: 5000, cat: "olives" },
      { id: 49, name: "زيتون أسود حب", price: 5000, cat: "olives" },
      { id: 50, name: "جبنة قطع حبة البركة", price: 9000, cat: "dairy" },
      { id: 51, name: "جبنة قطع نعنع", price: 9000, cat: "dairy" },
      { id: 52, name: "جبنة قطع زعتر", price: 9000, cat: "dairy" },
      { id: 53, name: "جبنة قطع سادة", price: 9000, cat: "dairy" },
      { id: 54, name: "لبنة خلطة بالشطة الحارة 600 غرام", price: 7000, cat: "dairy" },
      { id: 55, name: "لبنة خلطة بالزعتر 600 غرام", price: 7000, cat: "dairy" },
      { id: 56, name: "لبنة خلطة سادة بزيت الزيتون 600 غرام", price: 7000, cat: "dairy" },
      { id: 57, name: "لبنة خلطة بالبيتزا 600 غرام", price: 7000, cat: "dairy" },
      { id: 58, name: "لبنة بالمكسرات 600 غرام", price: 7000, cat: "dairy" },
      { id: 59, name: "لبنة كرات كوكتيل 600 غرام", price: 8000, cat: "dairy" },
      { id: 60, name: "لبنة كرات حبة البركة 600 غرام", price: 8000, cat: "dairy" },
      { id: 61, name: "لبنة كرات سادة 600 غرام", price: 8000, cat: "dairy" },
      { id: 62, name: "لبنة كرات بالزعتر 600 غرام", price: 8000, cat: "dairy" },
      { id: 63, name: "لبنة كرات بالنعنع 600 غرام", price: 8000, cat: "dairy" },
      { id: 64, name: "لبنة كرات بالشطة 600 غرام", price: 8000, cat: "dairy" }
    ];

    var productsNature = [
      { id: 100, name: "سيروم الشعر المغذي", price: 45000, cat: "hair" },
      { id: 101, name: "ماسك الذهب (Gold Mask)", price: 3000, cat: "masks" },
      { id: 102, name: "ماسك الألويفيرا (Aloe Vera Mask)", price: 3000, cat: "masks" },
      { id: 103, name: "ماسك الفحم (Charcoal Mask)", price: 3000, cat: "masks" },
      { id: 104, name: "غسول البشرة - روح الإنعاش", price: 2000, cat: "cleansers" },
      { id: 105, name: "كريم تجاعيد الوجه", price: 10000, cat: "creams" },
      { id: 106, name: "كريم حب الشباب", price: 10000, cat: "creams" },
      { id: 107, name: "كريم تفتيح البشرة", price: 10000, cat: "creams" },
      { id: 108, name: "كريم تفتيح البقع الداكنة", price: 12000, cat: "creams" },
      { id: 109, name: "كريم نهاري (Day Cream)", price: 15000, cat: "creams" },
      { id: 110, name: "كريم ليلي (Night Cream)", price: 15000, cat: "creams" },
      { id: 111, name: "كريم العناية بالأظافر", price: 12000, cat: "creams" },
      { id: 112, name: "كريم العناية بالقدم", price: 12000, cat: "creams" },
      { id: 113, name: "كريم خطوط الجلد", price: 12000, cat: "creams" },
      { id: 114, name: "لوشن للجسم (Body Lotion)", price: 8000, cat: "body" },
      { id: 115, name: "زيت عطري - جمال الطبيعة", price: 9000, cat: "oils" },
      { id: 116, name: "زيت عطري - عود", price: 9000, cat: "oils" },
      { id: 117, name: "زيت عطري - العطر الفرنسي", price: 9000, cat: "oils" },
      { id: 118, name: "زيت مساج (Massage Oil)", price: 5000, cat: "oils" },
      { id: 119, name: "سنفرة السكر - فراولة", price: 8000, cat: "scrubs" },
      { id: 120, name: "سنفرة السكر - جمال الطبيعة", price: 8000, cat: "scrubs" },
      { id: 121, name: "سنفرة السكر - أجمل إحساس", price: 8000, cat: "scrubs" },
      { id: 122, name: "سنفرة الوجه مع العسل", price: 12000, cat: "scrubs" },
      { id: 123, name: "زيت حشيشة الهليون الأخضر للشعر", price: 15000, cat: "hair" },
      { id: 124, name: "زيت تطويل الرموش", price: 5000, cat: "hair" },
      { id: 125, name: "شامبو خلاصة العسل والأعشاب", price: 8000, cat: "hair" },
      { id: 126, name: "صابون الشبة والمسك والبابونج", price: 3000, cat: "soap" },
      { id: 127, name: "صابون الحبة السوداء", price: 3000, cat: "soap" },
      { id: 128, name: "صابون زيت الأرغان وجنين القمح", price: 3000, cat: "soap" },
      { id: 129, name: "صابون الكركم والكبريت والعكبر", price: 3000, cat: "soap" }
    ];

    var catsJana = [
      { id: 'all', name: 'الكل' }, { id: 'dairy', name: 'الأجبان' }, { id: 'olives', name: 'الزيتون' },
      { id: 'drinks', name: 'الشرابات' }, { id: 'distillates', name: 'المقطرات' }, { id: 'pickles', name: 'المكدوس' },
      { id: 'spices', name: 'زعتر وسماق' }, { id: 'vinegar', name: 'الخل والدبس' }, { id: 'jams', name: 'المربيات' }, { id: 'oils', name: 'زيت الزيتون' }
    ];

    var catsNature = [
      { id: 'all', name: 'الكل' }, { id: 'masks', name: 'الماسكات' }, { id: 'cleansers', name: 'الغسولات' },
      { id: 'creams', name: 'الكريمات' }, { id: 'body', name: 'العناية والجسم' }, { id: 'oils', name: 'الزيوت والروائح' },
      { id: 'scrubs', name: 'السنفرات' }, { id: 'hair', name: 'الشعر والشامبو' }, { id: 'soap', name: 'الصابون الطبيعي' }
    ];

    var cart = [];
    var curCat = "all";

    function verifyPin() {
      var val = document.getElementById("pin-input").value;
      if (val === ADMIN_PIN) {
        isAdmin = true;
        document.getElementById("lock-screen").style.display = "none";
        applyRolePermissions();
        fetchStats();
      } else {
        var err = document.getElementById("lock-error");
        err.style.display = "block";
        setTimeout(() => err.style.display = "none", 2500);
      }
    }

    function enterAsCashier() {
      isAdmin = false;
      document.getElementById("lock-screen").style.display = "none";
      applyRolePermissions();
    }

    function lockSystem() {
      document.getElementById("pin-input").value = "";
      document.getElementById("lock-screen").style.display = "flex";
    }

    function applyRolePermissions() {
      var badge = document.getElementById("role-badge");
      var dlBtn = document.getElementById("dl-link");
      var statsBox = document.getElementById("stats-box");

      if (isAdmin) {
        badge.innerHTML = "👑 وضع: مدير (أدمن)";
        badge.style.background = "#dcfce7";
        badge.style.color = "#166534";
        badge.style.borderColor = "#86efac";
        dlBtn.style.display = "flex";
        statsBox.style.display = "flex";
      } else {
        badge.innerHTML = "🔒 وضع: كاشير (اضغط للتبديل)";
        badge.style.background = "#fef3c7";
        badge.style.color = "#d97706";
        badge.style.borderColor = "#fde68a";
        dlBtn.style.display = "none";
        statsBox.style.display = "none";
      }
    }

    function getActiveMeta() {
      if (currentCompany === 'nature') {
        return {
          name: "شركة جمال الطبيعة", subtitle: "نظام كاشير العناية والتجميل",
          primary: "#4a154b", leaf: "#e01a4f", products: productsNature, cats: catsNature,
          endpoint: '/save_nature', statsEndpoint: '/stats_nature', downloadLink: '/download_nature'
        };
      } else {
        return {
          name: "جنى الوديان اللبنانية", subtitle: "نظام الكاشير الموحد",
          primary: "#14361b", leaf: "#84b53a", products: productsJana, cats: catsJana,
          endpoint: '/save_jana', statsEndpoint: '/stats_jana', downloadLink: '/download_jana'
        };
      }
    }

    function switchCompany() {
      currentCompany = (currentCompany === 'jana') ? 'nature' : 'jana';
      localStorage.setItem('jw_active_company', currentCompany);
      cart = []; curCat = 'all';
      initTheme();
      renderCats();
      renderGrid();
      renderCart();
      if(isAdmin) fetchStats();
    }

    function initTheme() {
      var meta = getActiveMeta();
      document.documentElement.style.setProperty('--current-primary', meta.primary);
      document.documentElement.style.setProperty('--current-leaf', meta.leaf);
      document.getElementById('app-title').innerText = meta.name;
      document.getElementById('app-subtitle').innerText = meta.subtitle;
      document.getElementById('dl-link').href = meta.downloadLink;
      document.getElementById('dl-link').innerText = "📥 إكسل " + (currentCompany === 'nature' ? 'الجمال' : 'الجنى');
      document.querySelector('.btn-switch').innerText = (currentCompany === 'jana') ? "🔄 التبديل لجمال الطبيعة" : "🔄 التبديل لجنى الوديان";
      
      var percentBox = document.getElementById('percent-discounts-box');
      if (currentCompany === 'nature') {
        percentBox.style.display = 'flex';
      } else {
        percentBox.style.display = 'none';
        document.getElementById('disc').value = 0;
      }
    }

    function applyPercentDiscount(percent) {
      var sub = 0;
      cart.forEach(item => { sub += (item.price * item.qty); });
      var discountAmount = (sub * percent) / 100;
      document.getElementById('disc').value = discountAmount;
      renderCart();
    }

    function normalizeArabic(text) {
      return text.toString().replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').toLowerCase();
    }

    function onSearchInput() {
      var val = document.getElementById("search-input").value;
      document.getElementById("btn-clear-search").style.display = val.length > 0 ? "block" : "none";
      renderGrid();
    }

    function clearSearch() {
      document.getElementById("search-input").value = "";
      document.getElementById("btn-clear-search").style.display = "none";
      renderGrid();
    }

    function renderCats() {
      var meta = getActiveMeta();
      var container = document.getElementById('cats-container');
      container.innerHTML = '';
      meta.cats.forEach(function(c) {
        var btn = document.createElement('button');
        btn.className = 'cat-btn' + (curCat === c.id ? ' active' : '');
        btn.innerText = c.name;
        btn.onclick = function() { filterCat(c.id, btn); };
        container.appendChild(btn);
      });
    }

    function filterCat(c, btn) {
      curCat = c;
      document.querySelectorAll(".cat-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      renderGrid();
    }

    function renderGrid() {
      var meta = getActiveMeta();
      var g = document.getElementById("grid"); g.innerHTML = "";
      var q = normalizeArabic(document.getElementById("search-input").value.trim());
      var list = meta.products.filter(p => (curCat === "all" || p.cat === curCat) && (q === "" || normalizeArabic(p.name).includes(q) || p.price.toString().includes(q)));
      document.getElementById("search-results-count").innerText = `يعرض ${list.length} من ${meta.products.length} صنف`;
      
      if (list.length === 0) { 
        g.innerHTML = '<div style="grid-column: 1/-1; text-align:center; padding: 25px; background:#fff; border-radius:10px; color:#666;">لا توجد منتجات مطابقة..</div>'; 
        return; 
      }
      
      list.forEach(p => {
        var inCart = cart.find(i => i.id === p.id);
        var qtyBadge = inCart ? `<div class="card-badge" style="display:flex;">${inCart.qty}</div>` : `<div class="card-badge"></div>`;
        var card = document.createElement("div"); 
        card.className = "card";
        card.innerHTML = `${qtyBadge}<div><h4>${p.name}</h4><div class="price">${p.price.toLocaleString()} د.ع</div></div><button class="btn-add" onclick="addItem(${p.id})">+ إضافة للفاتورة</button>`;
        g.appendChild(card);
      });
    }

    function addItem(id) {
      var meta = getActiveMeta();
      var p = meta.products.find(x => x.id === id);
      if (!p) return;
      var found = cart.find(i => i.id === id);
      if (found) {
        found.qty += 1;
      } else { 
        cart.push({ id: p.id, name: p.name, price: p.price, qty: 1 }); 
      }
      renderCart(); renderGrid();
    }

    function updateQty(id, delta) {
      var item = cart.find(i => i.id === id);
      if (!item) return;
      item.qty += delta;
      if (item.qty <= 0) cart = cart.filter(i => i.id !== id);
      renderCart(); renderGrid();
    }

    function renderCart() {
      var list = document.getElementById("list");
      if (cart.length === 0) {
        list.innerHTML = '<p style="color:#888; text-align:center; margin:14px 0;">اضغط على المنتجات بالأعلى لإضافتها للفاتورة</p>';
        document.getElementById("sub").innerText = "0 د.ع"; document.getElementById("tot").innerText = "0 د.ع";
        document.getElementById("count").innerText = "0 قطع"; document.getElementById("float-tot").innerText = "0 د.ع"; document.getElementById("float-qty").innerText = "(0 قطع)";
        return;
      }
      list.innerHTML = ""; var sub = 0; var totalQty = 0;
      cart.forEach(item => {
        var lineTotal = item.price * item.qty; sub += lineTotal; totalQty += item.qty;
        var row = document.createElement("div"); row.className = "item";
        row.innerHTML = `<div><strong>${item.name}</strong><br><small style="color:#666;">${item.price.toLocaleString()} × ${item.qty} = <b>${lineTotal.toLocaleString()} د.ع</b></small></div><div style="display:flex; align-items:center; gap:8px;"><button class="q-btn" onclick="updateQty(${item.id}, -1)">-</button><span style="font-weight:bold;">${item.qty}</span><button class="q-btn" onclick="updateQty(${item.id}, 1)">+</button></div>`;
        list.appendChild(row);
      });
      var disc = parseFloat(document.getElementById("disc").value) || 0;
      var tot = Math.max(0, sub - disc);
      document.getElementById("sub").innerText = sub.toLocaleString() + " د.ع";
      document.getElementById("tot").innerText = tot.toLocaleString() + " د.ع";
      document.getElementById("count").innerText = totalQty + " قطع (" + cart.length + " صنف)";
      document.getElementById("float-tot").innerText = tot.toLocaleString() + " د.ع";
      document.getElementById("float-qty").innerText = "(" + totalQty + " قطع)";
    }

    function clearCart() { cart = []; document.getElementById("disc").value = 0; document.getElementById("notes").value = ""; renderCart(); renderGrid(); }
    function showToast(msg) { var t = document.getElementById("toast"); t.innerText = msg; t.style.display = "block"; setTimeout(() => { t.style.display = "none"; }, 3500); }

    // دالة طباعة الفاتورة
    function printReceipt() {
      if (cart.length === 0) { showToast("⚠️ السلة فارغة للطباعة!"); return; }
      var meta = getActiveMeta();
      document.getElementById("print-comp-name").innerText = meta.name;
      document.getElementById("print-date").innerText = new Date().toLocaleString('ar-IQ');
      document.getElementById("print-seller").innerText = document.getElementById("seller-name").value || "كاشير عام";
      
      var tbody = document.getElementById("print-items-body");
      tbody.innerHTML = "";
      var sub = 0;
      cart.forEach(i => {
        var lineTotal = i.price * i.qty;
        sub += lineTotal;
        tbody.innerHTML += `<tr><td style="padding:4px;">${i.name}</td><td style="text-align:center; padding:4px;">${i.qty}</td><td style="text-align:left; padding:4px;">${lineTotal.toLocaleString()}</td></tr>`;
      });
      var disc = parseFloat(document.getElementById("disc").value) || 0;
      var tot = Math.max(0, sub - disc);
      
      document.getElementById("print-sub").innerText = sub.toLocaleString() + " د.ع";
      document.getElementById("print-disc").innerText = disc.toLocaleString() + " د.ع";
      document.getElementById("print-tot").innerText = tot.toLocaleString() + " د.ع";
      
      window.print();
    }

    async function fetchStats() {
      if (!isAdmin) return;
      try {
        var meta = getActiveMeta();
        let res = await fetch(meta.statsEndpoint); let data = await res.json(); 
        document.getElementById('stat-total').innerText = 'إجمالي مبيعات الشركة: ' + data.total_sales.toLocaleString() + ' د.ع'; 
        document.getElementById('stat-count').innerText = 'عدد الفواتير: ' + data.count; 
      } catch(e) {}
    }

    async function submitOrder() {
      if (cart.length === 0) { showToast("⚠️ السلة فارغة!"); return; }
      var seller = document.getElementById("seller-name").value.trim() || "كاشير عام";
      var btn = document.getElementById("btn-save"); btn.disabled = true; btn.innerText = "جاري الحفظ...";
      
      var sub = 0; var totalQty = 0; var parts = [];
      cart.forEach(i => { sub += (i.price * i.qty); totalQty += i.qty; parts.push(i.name + " [×" + i.qty + "]"); });
      var disc = parseFloat(document.getElementById("disc").value) || 0;
      var payload = { 
        seller: seller, 
        itemsText: parts.join(" | "), totalQty: totalQty, subtotal: sub, 
        discount: disc, total: Math.max(0, sub - disc), notes: document.getElementById("notes").value 
      };
      
      try {
        var meta = getActiveMeta();
        let res = await fetch(meta.endpoint, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
        let data = await res.json();
        if(data.status === 'ok') {
          showToast("✅ تم حفظ الفاتورة بنجاح!");
          clearCart(); if(isAdmin) fetchStats();
        } else { showToast("⚠️ خطأ في الحفظ"); }
      } catch (err) { alert("خطأ في الاتصال بالسيرفر."); }
      btn.disabled = false; btn.innerText = "💾 حفظ الفاتورة";
    }

    initTheme();
    renderCats();
    renderGrid();
  </script>
</body>
</html>
"""

@app.route('/')
def index():
    return Response(HTML_PAGE, mimetype='text/html')

@app.route('/stats_jana')
def stats_jana():
    with file_lock:
        df = pd.read_excel(EXCEL_JANA)
        return jsonify({'total_sales': int(df['الصافي المستلم (د.ع)'].sum()) if len(df) > 0 else 0, 'count': len(df)})

@app.route('/save_jana', methods=['POST'])
def save_jana():
    data = request.json
    baghdad_tz = timezone(timedelta(hours=3))
    now_str = datetime.now(baghdad_tz).strftime('%Y-%m-%d %H:%M:%S')
    inv_id = 'JW-' + datetime.now(baghdad_tz).strftime('%H%M%S')
    with file_lock:
        new_row = {'رقم الفاتورة': inv_id, 'التاريخ والوقت': now_str, 'اسم البائع / الكاشير': data.get('seller', 'كاشير عام'), 'تفاصيل المواد المباعة': data['itemsText'], 'عدد القطع': data.get('totalQty', 1), 'المجموع الفرعي (د.ع)': data['subtotal'], 'الخصم (د.ع)': data['discount'], 'الصافي المستلم (د.ع)': data['total'], 'ملاحظات': data.get('notes') or '-'}
        df_sales = pd.read_excel(EXCEL_JANA)
        pd.concat([df_sales, pd.DataFrame([new_row])], ignore_index=True).to_excel(EXCEL_JANA, index=False)
    return jsonify({'status': 'ok', 'invoice_id': inv_id})

@app.route('/download_jana')
def download_jana():
    with file_lock: return send_file(EXCEL_JANA, as_attachment=True, download_name='مبيعات_جنى_الوديان_الموحد.xlsx')

@app.route('/stats_nature')
def stats_nature():
    with file_lock:
        df = pd.read_excel(EXCEL_NATURE)
        return jsonify({'total_sales': int(df['الصافي المستلم (د.ع)'].sum()) if len(df) > 0 else 0, 'count': len(df)})

@app.route('/save_nature', methods=['POST'])
def save_nature():
    data = request.json
    baghdad_tz = timezone(timedelta(hours=3))
    now_str = datetime.now(baghdad_tz).strftime('%Y-%m-%d %H:%M:%S')
    inv_id = 'NAT-' + datetime.now(baghdad_tz).strftime('%H%M%S')
    with file_lock:
        new_row = {'رقم الفاتورة': inv_id, 'التاريخ والوقت': now_str, 'اسم البائع / الكاشير': data.get('seller', 'كاشير عام'), 'تفاصيل المواد المباعة': data['itemsText'], 'عدد القطع': data.get('totalQty', 1), 'المجموع الفرعي (د.ع)': data['subtotal'], 'الخصم (د.ع)': data['discount'], 'الصافي المستلم (د.ع)': data['total'], 'ملاحظات': data.get('notes') or '-'}
        df_sales = pd.read_excel(EXCEL_NATURE)
        pd.concat([df_sales, pd.DataFrame([new_row])], ignore_index=True).to_excel(EXCEL_NATURE, index=False)
    return jsonify({'status': 'ok', 'invoice_id': inv_id})

@app.route('/download_nature')
def download_nature():
    with file_lock: return send_file(EXCEL_NATURE, as_attachment=True, download_name='مبيعات_جمال_الطبيعة.xlsx')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
