// Available groups will be populated at runtime
const availableGroups = [];
let currentGroup = '';

// Search functionality
let searchResults = [];
let currentResultIndex = -1;

// Initialize the page
window.addEventListener('DOMContentLoaded', async () => {
  // Apply saved theme preference
  applyTheme();
  
  // Set up theme toggle functionality
  document.getElementById('theme-toggle').addEventListener('change', toggleTheme);
  
  // Set up search functionality
  document.getElementById('search-button').addEventListener('click', performSearch);
  document.getElementById('search-input').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
      performSearch();
    }
  });
  document.getElementById('next-result').addEventListener('click', navigateToNextResult);
  document.getElementById('prev-result').addEventListener('click', navigateToPrevResult);
  
  // Set up back to top button and sticky header
  setupScrollFunctionality();
  
  await loadAvailableGroups();
  
  // Get group from URL if present
  const urlParams = new URLSearchParams(window.location.search);
  const groupParam = urlParams.get('group');
  
  if (groupParam && availableGroups.includes(groupParam)) {
    currentGroup = groupParam;
    highlightActiveGroup(currentGroup);
    loadMessages(currentGroup);
  }
});

// Theme management functions
function applyTheme() {
  const savedTheme = localStorage.getItem('theme') || 'light';
  document.documentElement.setAttribute('data-theme', savedTheme);
  document.getElementById('theme-toggle').checked = savedTheme === 'dark';
}

function toggleTheme() {
  const isDarkMode = document.getElementById('theme-toggle').checked;
  const theme = isDarkMode ? 'dark' : 'light';
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('theme', theme);
}

// Function to discover available groups
async function loadAvailableGroups() {
  try {
    // Try to load the groups configuration file
    const configResponse = await fetch('./groups-config.json');
    if (configResponse.ok) {
      const config = await configResponse.json();
      if (Array.isArray(config.groups)) {
        availableGroups.push(...config.groups);
      }
    }
    
    // If no groups were loaded from config, try fallback methods
    if (availableGroups.length === 0) {
      // Fallback to hardcoded groups or directory scanning would go here
      console.warn("No groups found in config file");
    }
    
    // Sort groups alphabetically
    availableGroups.sort();
    
    // Populate sidebar navigation
    const navElement = document.getElementById('group-nav');
    navElement.innerHTML = '';
    
    if (availableGroups.length === 0) {
      navElement.innerHTML = '<p class="error">No groups found</p>';
      return;
    }
    
    // Add each group as a nav item
    availableGroups.forEach(group => {
      const navItem = document.createElement('div');
      navItem.classList.add('nav-item');
      navItem.textContent = group;
      navItem.dataset.group = group;
      navItem.addEventListener('click', () => {
        currentGroup = group;
        
        // Update URL without reloading the page
        const url = new URL(window.location);
        url.searchParams.set('group', currentGroup);
        window.history.pushState({}, '', url);
        
        // Update UI
        highlightActiveGroup(currentGroup);
        loadMessages(currentGroup);
      });
      navElement.appendChild(navItem);
    });
  } catch (error) {
    console.error('Error loading available groups:', error);
    document.getElementById('nav-loading').textContent = 'Error loading groups';
  }
}

// Function to highlight the active group in the sidebar
function highlightActiveGroup(groupName) {
  // Remove active class from all nav items
  document.querySelectorAll('.nav-item').forEach(item => {
    item.classList.remove('active');
  });
  
  // Add active class to the selected group
  const activeItem = document.querySelector(`.nav-item[data-group="${groupName}"]`);
  if (activeItem) {
    activeItem.classList.add('active');
    // Ensure the active item is visible (scroll to it if needed)
    activeItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
}

// Reusable function to load JSON data from a file
async function loadJsonFile(fileName) {
  const response = await fetch(fileName);
  if (!response.ok) {
    throw new Error(`HTTP error! Status: ${response.status}`);
  }
  return await response.json();
}

// Function to load and process the messages
async function loadMessages(groupName) {
  if (!groupName) {
    document.getElementById('loading').textContent = 'Please select a group to view messages';
    document.getElementById('page-title').textContent = 'Dynamics 365 and Power Platform Preview Programs';
    document.getElementById('message-board').innerHTML = '<p id="loading">Please select a group to view messages</p>';
    return;
  }
  
  document.getElementById('message-board').innerHTML = '<p id="loading">Loading messages...</p>';
  document.getElementById('page-title').textContent = groupName;
  
  try {
    // Update the path to include the group subdirectory
    const data = await loadJsonFile(`${groupName}/${groupName} Messages.json`);
    const messages = data.body.value;
    
    // Also load references for user information
    let userMap = {};
    try {
      // Update the path to include the group subdirectory
      const refData = await loadJsonFile(`${groupName}/${groupName} References.json`);
      // Create a map of user IDs to names if the references contain user data
      if (refData && refData.body && refData.body.value) {
        userMap = refData.body.value.reduce((map, ref) => {
          if (ref.id && ref.full_name) {
            map[ref.id] = ref.full_name;
          }
          return map;
        }, {});
      }
    } catch (refError) {
      console.warn('Could not load user references:', refError);
    }
    
    // Build message map
    const messageMap = {};
    messages.forEach(msg => {
      // Skip entries that don't have an ID or are just placeholders
      if (!msg.id || (Object.keys(msg).length === 1 && msg.group_created_id)) {
        return;
      }
      msg.children = [];
      messageMap[msg.id] = msg;
    });

    // Sort all messages by created_at in ascending order (oldest first)
    messages.sort((a, b) => new Date(a.created_at) - new Date(b.created_at));

    // Build hierarchy
    const roots = [];
    let missingParentCount = 0; // Track messages with missing parents
    
    messages.forEach(msg => {
      // Skip entries that don't have an ID or are just placeholders
      if (!msg.id || (Object.keys(msg).length === 1 && msg.group_created_id)) {
        return;
      }
      
      if (msg.replied_to_id) {
        if (messageMap[msg.replied_to_id]) {
          messageMap[msg.replied_to_id].children.push(msg);
        } else {
          roots.push(msg); // If parent is missing, treat as root
          missingParentCount++; // Increment counter for missing parents
        }
      } else {
        roots.push(msg);
      }
    });

    // Sort root messages and replies by date
    roots.sort((a, b) => new Date(b.created_at) - new Date(a.created_at));
    Object.values(messageMap).forEach(msg => {
      msg.children.sort((a, b) => new Date(a.created_at) - new Date(b.created_at));
    });

    // Remove loading message and update the message board
    document.getElementById('message-board').innerHTML = '';
    
    // Add message count summary
    const summaryElement = document.createElement('p');
    summaryElement.classList.add('summary');
    summaryElement.textContent = `Displaying ${roots.length} threads with ${messages.length} total messages`;
    
    // Add missing parent info if any exist
    if (missingParentCount > 0) {
      summaryElement.textContent += ` (${missingParentCount} messages with missing parents)`;
    }
    
    document.getElementById('message-board').appendChild(summaryElement);
    
    // Render and attach messages
    document.getElementById('message-board').appendChild(renderMessages(roots, userMap));
  } catch (error) {
    console.error(`Error loading messages for group ${groupName}:`, error);
    const errorMsg = document.createElement('p');
    errorMsg.classList.add('error');
    errorMsg.textContent = `Failed to load messages for ${groupName}: ${error.message}`;
    document.getElementById('loading').replaceWith(errorMsg);
  }
}

// Helper function to get user name from ID
function getUserName(userId, userMap) {
  return userMap[userId] || `User ID: ${userId}`;
}

// Improved rendering with formatted date and message styling
function renderMessages(messages, userMap = {}) {
  const ul = document.createElement('ul');
  messages.forEach(msg => {
    const li = document.createElement('li');
    li.classList.add('message');
    
    // Create message header with sender and date
    const header = document.createElement('div');
    header.classList.add('message-header');
    
    const sender = document.createElement('span');
    sender.classList.add('sender');
    sender.textContent = getUserName(msg.sender_id, userMap);
    header.appendChild(sender);
    
    const date = document.createElement('span');
    date.classList.add('date');
    // Format the date nicely
    if (msg.created_at) {
      try {
        // Parse the date string more reliably
        const dateString = msg.created_at.replace(/(\d{4})\/(\d{2})\/(\d{2})\s(\d{2}):(\d{2}):(\d{2})\s\+(\d{4})/, 
          '$1-$2-$3T$4:$5:$6+$7:00');
        const dateObj = new Date(dateString);
        
        // Check if the date is valid before using it
        if (!isNaN(dateObj.getTime())) {
          date.textContent = dateObj.toLocaleString();
        } else {
          date.textContent = msg.created_at; // Fall back to the original string
        }
      } catch (e) {
        date.textContent = msg.created_at; // Fall back to the original string
      }
    } else {
      date.textContent = 'No date'; // Handle missing date
    }
    header.appendChild(date);
    
    li.appendChild(header);
    
    // Create message content
    const content = document.createElement('div');
    content.classList.add('message-content');
    
    // Use rich content if available, otherwise use plain or parsed
    if (msg.body?.rich) {
      content.innerHTML = msg.body.rich;
    } else if (msg.body?.plain) {
      content.textContent = msg.body.plain;
    } else if (msg.body?.parsed) {
      content.textContent = msg.body.parsed;
    } else if (msg.content_excerpt) {
      content.textContent = msg.content_excerpt;
    } else {
      content.textContent = '[No content]';
    }
    
    li.appendChild(content);

    if (msg.children.length > 0) {
      const toggle = document.createElement('span');
      toggle.textContent = `Show Replies (${msg.children.length})`;
      toggle.classList.add('toggle');
      let shown = false;
      const childrenUl = renderMessages(msg.children, userMap);
      childrenUl.style.display = 'none';
      toggle.onclick = () => {
        shown = !shown;
        childrenUl.style.display = shown ? 'block' : 'none';
        toggle.textContent = shown ? `Hide Replies (${msg.children.length})` : `Show Replies (${msg.children.length})`;
      };
      li.appendChild(toggle);
      li.appendChild(childrenUl);
    }

    ul.appendChild(li);
  });
  return ul;
}

// Perform search across all loaded messages
function performSearch() {
  // Reset previous search
  clearSearch();
  
  const searchTerm = document.getElementById('search-input').value.trim().toLowerCase();
  if (!searchTerm) return;
  
  // Find all message content elements
  const messageElements = document.querySelectorAll('.message-content');
  if (messageElements.length === 0) {
    document.getElementById('search-count').textContent = 'No messages to search';
    return;
  }
  
  // Search through messages
  messageElements.forEach(element => {
    const originalText = element.innerHTML;
    const lowerText = element.textContent.toLowerCase();
    
    // If we found a match
    if (lowerText.includes(searchTerm)) {
      // Add to search results array
      searchResults.push(element);
      
      // Highlight the search term
      const regex = new RegExp(`(${escapeRegExp(searchTerm)})`, 'gi');
      element.innerHTML = originalText.replace(regex, '<span class="search-highlight">$1</span>');
    }
  });
  
  // Update UI with search results
  updateSearchUI();
  
  // Navigate to first result
  if (searchResults.length > 0) {
    navigateToResult(0);
  }
}

// Clear previous search results
function clearSearch() {
  // Clear any existing highlights
  document.querySelectorAll('.search-highlight, .search-active').forEach(el => {
    const parent = el.parentNode;
    parent.replaceChild(document.createTextNode(el.textContent), el);
    parent.normalize(); // Combine adjacent text nodes
  });
  
  // Reset search state
  searchResults = [];
  currentResultIndex = -1;
  
  // Reset UI
  document.getElementById('search-count').textContent = '';
  document.getElementById('prev-result').disabled = true;
  document.getElementById('next-result').disabled = true;
}

// Update search UI elements
function updateSearchUI() {
  const count = searchResults.length;
  document.getElementById('search-count').textContent = 
    count > 0 ? `${count} result${count !== 1 ? 's' : ''} found` : 'No results found';
  
  document.getElementById('prev-result').disabled = count === 0 || currentResultIndex <= 0;
  document.getElementById('next-result').disabled = count === 0 || currentResultIndex >= count - 1;
}

// Navigate to a specific search result
function navigateToResult(index) {
  if (index < 0 || index >= searchResults.length) return;
  
  // Remove active highlight from previous result
  document.querySelectorAll('.search-active').forEach(el => {
    el.classList.replace('search-active', 'search-highlight');
  });
  
  // Set new active result
  currentResultIndex = index;
  
  const currentResult = searchResults[currentResultIndex];
  
  // Add active highlight to current result
  const highlights = currentResult.querySelectorAll('.search-highlight');
  if (highlights.length > 0) {
    highlights[0].classList.replace('search-highlight', 'search-active');
  }
  
  // Scroll the result into view
  currentResult.closest('.message').scrollIntoView({
    behavior: 'smooth',
    block: 'center'
  });
  
  // Update navigation buttons
  updateSearchUI();
}

// Navigate to next search result
function navigateToNextResult() {
  if (currentResultIndex < searchResults.length - 1) {
    navigateToResult(currentResultIndex + 1);
  }
}

// Navigate to previous search result
function navigateToPrevResult() {
  if (currentResultIndex > 0) {
    navigateToResult(currentResultIndex - 1);
  }
}

// Helper function to escape special regex characters in search term
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Set up scroll functionality including back to top button and sticky header
function setupScrollFunctionality() {
  const backToTopBtn = document.getElementById('back-to-top');
  const header = document.querySelector('header');
  
  // Show/hide button based on scroll position and add sticky header class
  window.addEventListener('scroll', () => {
    const scrollPosition = document.body.scrollTop || document.documentElement.scrollTop;
    
    // Handle back to top button
    if (scrollPosition > 300) {
      backToTopBtn.style.display = 'block';
    } else {
      backToTopBtn.style.display = 'none';
    }
    
    // Handle sticky header
    if (scrollPosition > 0) {
      document.body.classList.add('sticky-header');
    } else {
      document.body.classList.remove('sticky-header');
    }
  });
  
  // Scroll to top when clicked
  backToTopBtn.addEventListener('click', () => {
    document.body.scrollTop = 0; // For Safari
    document.documentElement.scrollTop = 0; // For Chrome, Firefox, IE and Opera
  });
}

// Start loading the messages
loadMessages();
