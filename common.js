// Shared Common Utilities for Thread Track

// Check Authentication Session
(function checkSession() {
  const currentPath = window.location.pathname;
  const isLoginPage = currentPath.endsWith('login.html');
  const userSession = localStorage.getItem('user_session');

  if (!userSession && !isLoginPage) {
    window.location.href = 'login.html';
  } else if (userSession && isLoginPage) {
    window.location.href = 'index.html';
  }
})();

// Get Active User Profile
function getSessionUser() {
  const sessionStr = localStorage.getItem('user_session');
  return sessionStr ? JSON.parse(sessionStr) : null;
}

// Log Out
function logout() {
  localStorage.removeItem('user_session');
  window.location.href = 'login.html';
}

// Global Theme Management
function initTheme() {
  const currentTheme = localStorage.getItem('theme') || 'light';
  document.documentElement.setAttribute('data-theme', currentTheme);
  updateThemeIcon(currentTheme);
}

function toggleTheme() {
  const currentTheme = document.documentElement.getAttribute('data-theme');
  const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', newTheme);
  localStorage.setItem('theme', newTheme);
  updateThemeIcon(newTheme);
}

function updateThemeIcon(theme) {
  const icon = document.getElementById('theme-icon');
  if (icon) {
    icon.className = theme === 'dark' ? 'bi bi-sun-fill' : 'bi bi-moon-fill';
  }
}

// Call Apps Script API
async function callApi(action, params = {}, isPost = false, body = null) {
  const user = getSessionUser();
  const spreadsheetId = user ? user.spreadsheetId : null;
  
  const url = new URL(CONFIG.APPS_SCRIPT_URL);
  
  if (!isPost) {
    // GET request
    url.searchParams.append('action', action);
    if (spreadsheetId) {
      url.searchParams.append('spreadsheetId', spreadsheetId);
    }
    for (const key in params) {
      url.searchParams.append(key, params[key]);
    }
  }

  const options = {
    method: isPost ? 'POST' : 'GET',
    mode: 'cors'
  };

  if (isPost) {
    const postPayload = {
      action: action,
      spreadsheetId: spreadsheetId,
      ...body
    };
    options.body = JSON.stringify(postPayload);
    options.headers = {
      'Content-Type': 'text/plain;charset=utf-8' // Crucial for Apps Script CORS and POST bodies
    };
  }

  try {
    const response = await fetch(url.toString(), options);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const result = await response.json();
    if (!result.success) {
      throw new Error(result.error || 'Operation failed');
    }
    return result;
  } catch (error) {
    console.error(`API Error on ${action}:`, error);
    showAlert(error.message || 'Server communication error', 'danger');
    throw error;
  }
}

// Display Alerts
function showAlert(message, type = 'success', containerId = 'alert-container') {
  const container = document.getElementById(containerId);
  if (!container) return;

  const alertHtml = `
    <div class="alert alert-${type} alert-dismissible fade show" role="alert">
      ${message}
      <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
    </div>
  `;
  container.innerHTML = alertHtml;
  
  // Auto-dismiss after 5 seconds if not danger
  if (type !== 'danger') {
    setTimeout(() => {
      const alertEl = container.querySelector('.alert');
      if (alertEl) {
        const bsAlert = new bootstrap.Alert(alertEl);
        bsAlert.close();
      }
    }, 5000);
  }
}

// Helpers
function formatCurrency(value) {
  return '₹' + parseFloat(value || 0).toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function formatDate(dateStr) {
  if (!dateStr) return 'Pending';
  const d = new Date(dateStr);
  return d.toLocaleDateString('en-IN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}

// Inject Shared Navigation Bar
function renderNavbar() {
  const container = document.getElementById('navbar-placeholder');
  if (!container) return;

  const user = getSessionUser();
  if (!user) return;

  const isAdmin = user.role === 'Admin';
  const currentPath = window.location.pathname;
  
  const navItems = [
    { name: 'Dashboard', path: 'index.html', icon: 'bi bi-grid-fill' },
    { name: 'Saree Lots', path: 'sarees.html', icon: 'bi bi-layers-fill' },
    { name: 'Vendors', path: 'vendors.html', icon: 'bi bi-people-fill' },
    { name: 'Payments', path: 'payments.html', icon: 'bi bi-cash-stack' },
    { name: 'Reports', path: 'reports.html', icon: 'bi bi-file-earmark-bar-graph-fill' }
  ];

  if (isAdmin) {
    navItems.push({ name: 'Admin Panel', path: 'admin.html', icon: 'bi bi-shield-lock-fill' });
  }

  const linksHtml = navItems.map(item => {
    const isActive = currentPath.endsWith(item.path) || (item.path === 'index.html' && currentPath.endsWith('/'));
    return `
      <li class="nav-item">
        <a class="nav-link nav-link-custom ${isActive ? 'active' : ''}" href="${item.path}">
          <i class="${item.icon} me-2"></i>${item.name}
        </a>
      </li>
    `;
  }).join('');

  const navbarHtml = `
    <nav class="navbar navbar-expand-lg navbar-custom">
      <div class="container">
        <a class="navbar-brand d-flex align-items-center" href="index.html">
          <i class="bi bi-activity me-2"></i>THREAD TRACK
        </a>
        <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarNav" aria-controls="navbarNav" aria-expanded="false" aria-label="Toggle navigation">
          <span class="navbar-toggler-icon"></span>
        </button>
        <div class="collapse navbar-collapse" id="navbarNav">
          <ul class="navbar-nav me-auto mb-2 mb-lg-0 gap-1">
            ${linksHtml}
          </ul>
          <div class="d-flex align-items-center gap-3">
            <button class="theme-toggle-btn" onclick="toggleTheme()" title="Toggle Dark/Light Mode">
              <i id="theme-icon" class="bi bi-moon-fill"></i>
            </button>
            <div class="dropdown">
              <button class="btn btn-outline-secondary dropdown-toggle d-flex align-items-center gap-2" type="button" id="userDropdown" data-bs-toggle="dropdown" aria-expanded="false" style="border-radius: var(--radius-sm);">
                <i class="bi bi-person-circle"></i>
                <span>${user.name}</span>
              </button>
              <ul class="dropdown-menu dropdown-menu-end" aria-labelledby="userDropdown" style="border-radius: var(--radius-sm); border: 1px solid var(--border-color); box-shadow: var(--shadow-md);">
                <li><a class="dropdown-item text-danger d-flex align-items-center gap-2" href="#" onclick="logout(); return false;"><i class="bi bi-box-arrow-right"></i> Logout</a></li>
              </ul>
            </div>
          </div>
        </div>
      </div>
    </nav>
  `;
  container.innerHTML = navbarHtml;
  initTheme();
}

// Initialise navbar and theme load on DOMContentLoaded
document.addEventListener('DOMContentLoaded', () => {
  renderNavbar();
});
