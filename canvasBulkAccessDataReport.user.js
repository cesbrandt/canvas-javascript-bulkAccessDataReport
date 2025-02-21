// ==UserScript==
// @name         Canvas Bulk Access Report Data
// @version      1.0
// @description  Generates a .CSV download of the access report for all students in a course
// @match        https://*.instructure.com/accounts/*/terms
// @match        https://*.instructure.com/accounts/*/terms?*
// @match        https://*.instructure.com/courses/*/users
// @updateURL     https://raw.githubusercontent.com/cesbrandt/canvas-javascript-bulkAccessDataReport/master/canvasBulkAccessDataReport.user.js
// ==/UserScript==

const delay = ms => new Promise(res => setTimeout(res, ms));

Node.prototype.getElementByInnerTextQuerySelector = function(value, selector, exact = false) {
	if(typeof value == 'string' || typeof value == 'number') {
		selector = (selector == null || selector == undefined) ? '*' : selector;
		var trimmedValue = String(value).trim();
		var dom = this.querySelectorAll(selector);
		for(var i = 0; i < dom.length; i++) {
			if(typeof dom[i] === 'object' && dom[i].innerText !== undefined) {
				var text = dom[i].innerText.trim();
				if(exact ? text === trimmedValue : text.match(new RegExp(trimmedValue.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'))) {
					return dom[i];
				}
			}
		}

		return undefined;
	}

	throw new TypeError("value must be a string or number, " + typeof value + " provided");
};

/**
 * @name          Build Variable Object
 * @description   Creates an object of GET variables and their values from a supplied URL
 * @return obj    Object of GET variables and their values
 */
String.prototype.buildVarObj = function() {
	var varObj = {};
	var vars = this.split('?');

	if(vars.length > 1) {
		vars = vars[1].split('&');
		for(var i in vars) {
			vars[i] = vars[i].split('=');
			varObj[vars[i][0]] = vars[i][1];
		}
	}

	return varObj;
};

/**
 * @name          Is Object or Array Empty?
 * @description   Generic function for determing if a JavaScript object or array is empty
 * @return undefined
 */
let isEmpty = obj => {
	if(Object.prototype.toString.call(obj) == '[object Array]') {
		return obj.length > 0 ? false : true;
	} else {
		for(var key in obj) {
			if(obj.hasOwnProperty(key)) {
				return false;
			}
		}

		return true;
	}
};

/**
 * @name          Extend Arrays/Objects
 * @description   Extends two arrays/objects
 * @return array/object
 */
let extend = (one, two) => {
	var extended;

	if(Object.prototype.toString.call(one) == '[object Object]' || Object.prototype.toString.call(two) == '[object Object]') {
		extended = {};
		for(var key in one) {
			extended[key] = one[key];
		}
		for(key in two) {
			extended[key] = two[key];
		}
	} else {
		extended = one;
		var i = extended.length;
		for(var j in two) {
			extended[i++] = two[j];
		}
	}

	return extended;
};

/**
 * @name          Data-to-HTTP
 * @description   Converts an object to HTTP parameters
 * @return string
 */
let dataToHttp = obj => {
	var pairs = [];

	for(var prop in obj) {
		if(!obj.hasOwnProperty(prop)) {
			continue;
		}
		if(Object.prototype.toString.call(obj[prop]) == '[object Object]') {
			pairs.push(dataToHttp(obj[prop]));
			continue;
		}
		pairs.push(prop + '=' + obj[prop]);
	}

	return pairs.join('&');
};

/**
 * @name          Format Cookie Name
 * @description   Returns cookie name formatted for site
 * @return string
 */
let formatCookieName = name => {
	return url.match(/(?!\/\/)[a-zA-Z1-3]*(?=\.)/) + '_' + view + (viewID !== null ? '_' + viewID : '') + '_' + escape(name);
};

/**
 * @name          Get Cookie by Name
 * @description   Returns cookie value
 * @return string
 */
let getCookie = (name, format) => {
	let value = document.cookie.match('(^|[^;]+)\\s*' + (typeof format !== 'undefined' && format ? formatCookieName(name) : name) + '\\s*=\\s*([^;]+)');

	return value !== null && value !== '' ? decodeURIComponent(value.pop()) : null;
};

let deepMergeObject = (targetObject = {}, sourceObject = {}) => {
	var copyTargetObject = JSON.parse(JSON.stringify(targetObject));
	var copySourceObject = JSON.parse(JSON.stringify(sourceObject));

	if(Array.isArray(copySourceObject) && Array.isArray(copyTargetObject)) {
		copyTargetObject = [...copyTargetObject, ...copySourceObject];
	} else {
		Object.keys(copySourceObject).forEach((key) => {
			if(Array.isArray(copySourceObject[key]) && Array.isArray(copyTargetObject[key])) {
				copyTargetObject[key] = [...copyTargetObject[key], ...copySourceObject[key]];
			} else if(typeof copySourceObject[key] === "object" && !Array.isArray(copySourceObject[key])) {
				copyTargetObject[key] = deepMergeObject(copyTargetObject[key], copySourceObject[key]);
			} else {
				copyTargetObject[key] = copySourceObject[key];
			}
		});
	}

	return copyTargetObject;
};

/**
 * @name          API Call
 * @description   Calls the Canvas API, handles multithreading, pagination, throttling
 * @param {string} type - GET, POST, DELETE, etc.
 * @param {string|array} context - URL path or array of path segments
 * @param {number|string} page - Page number (numeric or alphanumeric)
 * @param {array} getVars - Query parameters for the API call
 * @param {number|string} lastPage - Last page in pagination
 * @param {array} firstCall - Results of the previous calls (for merging)
 * @param {boolean} next - Indicates if there are more pages to fetch
 * @param {number} threads - Optional: number of concurrent threads (default: 5)
 * @returns {Promise<object>} - Merged result of all API calls across pages
 */
let callAPI = async (type, context, page = 1, getVars = [{}], threads = 5, lastPage = -1, firstCall = [{}], next = false) => {
	if(!window.bulkAccessRunning) {
		return {};
	}

  // Default `threads` to 5 if non-numeric or null, cap at 20
  threads = isNaN(threads) || threads === null || threads < 1 ? 5 : Math.min(threads, 20);

  // Build URL from context (either string or array)
  context = Array.isArray(context) ? '/api/v1/' + context.join('/') : context;
  let callURL = url.split('/' + view + '/')[0] + context;

  getVars = getVars || [{}]; // Default getVars if null
  getVars[0].page = page;
  getVars[0].per_page = 100;

  // First call to fetch the initial data and headers
  let initialResponse = await callAJAX(type, callURL, getVars[0]);
  if(initialResponse.status !== 200) {
    throw new Error('There was an error. Please try again.');
  }

  let json = JSON.parse(initialResponse.data.replace('while(1);', ''));
  let results = (firstCall.length === 1 && isEmpty(firstCall[0])) ? json : deepMergeObject(firstCall, json);

  // Function to fetch a single page of data
  const fetchPage = async (page) => {
    getVars[0].page = page;
    const response = await callAJAX(type, callURL, getVars[0]);
    if (response.status !== 200) {
      throw new Error('Error fetching page ' + page);
    }
    return response;
  };

  // Handle pagination via the "Link" header
  let linkHeader = initialResponse.headers['link'];
  let pages = parsePaginationLinks(linkHeader);
  lastPage = pages.last != undefined ? pages.last : lastPage;

  if(pages.next || lastPage > page) {
    next = true; // Ensure we fetch the next pages
    lastPage = pages.last || lastPage; // Get last page if known
  }

  var lowestThreshold = 700;

  // Adjust logic for scenarios
  if(next) {
    if(!isNaN(lastPage) && lastPage > -1) {
      // Scenario 1: Use multithreading when lastPage is numeric and not default
      let promises = [];
      let currentPage = page + 1;

      // Launch multiple threads to fetch pages concurrently
      for(let i = 0; i < threads && currentPage <= lastPage; i++) {
        promises.push(fetchPageInParallel(type, callURL, currentPage++, getVars));
      }

      // Await all thread promises to resolve and merge the results
      let pageResults = await Promise.all(promises);
      pageResults.forEach(pageData => {
        results = deepMergeObject(results, pageData);
      });
    } else if(!isNaN(lastPage)) {
      // Scenario 3: Use multithreading when lastPage is numeric and default
      let currentPage = page + 1;

      // Dynamic multithreading as we discover `lastPage`
      let isFetching = false;
      while(true) {
        var promises = [];

        // Launch concurrent requests in batches based on the threads count
        for(let i = 0; i < threads; i++) {
          promises.push(fetchPage(currentPage++));
        }

        // Wait for the batch of promises to complete
        isFetching = true;
        var pagesData = await Promise.all(promises);
        isFetching = false;

        // Process each page's data
        for(var pageData of pagesData) {
          results = deepMergeObject(results, JSON.parse(pageData.data.replace('while(1);', '')));
          lowestThreshold = pageData.headers['x-rate-limit-remaining'] < lowestThreshold ? pageData.headers['x-rate-limit-remaining'] : lowestThreshold;
        }

        // Check the `Link` header for the next/last page
        var lastPageResponse = pagesData[pagesData.length - 1]; // Get the last page from the current batch
        linkHeader = lastPageResponse.headers?.link;
        pages = parsePaginationLinks(linkHeader);

        // If no more `next` link or lastPage is reached, exit loop
        if(!pages.next) {
          break;
        }
      }
    } else {
      // Scenarios 2 and 4: No multithreading, iterative fetching based on `next`
      while(next && pages.next) {
        page++;
        let nextResponse = await fetchPageInParallel(type, callURL, page, getVars);
        results = deepMergeObject(results, nextResponse);

        let nextLinkHeader = nextResponse.headers['link'];
        pages = parsePaginationLinks(nextLinkHeader);
        next = !!pages.next;
      }
    }
  }

  return results;
};

/**
 * @name fetchPageInParallel
 * @description Helper function to fetch a specific page in parallel
 * @param {string} type - GET, POST, etc.
 * @param {string} callURL - URL to fetch
 * @param {number|string} page - Page to fetch
 * @param {object} getVars - Query parameters for the API call
 * @returns {Promise<object>} - JSON data of the fetched page
 */
let fetchPageInParallel = async (type, callURL, page, getVars) => {
	try {
		getVars[0].page = page;

		var response = await callAJAX(type, callURL, getVars[0]);
		if(response.status !== 200) {
			throw new Error('Error fetching page ' + page);
		}

		var json = JSON.parse(response.data.replace('while(1);', ''));

		return json;
	} catch(error) {
		console.error(`Error fetching page ${page}: ${error.message}`);

		return { error: error.message };
	}
};

/**
 * @name parsePaginationLinks
 * @description Parse the pagination links from the API response headers
 * @param {string} linkHeader - The 'Link' header from the API response
 * @returns {object} - Parsed links for pagination
 */
let parsePaginationLinks = linkHeader => {
	if(!linkHeader) {
		return {};
	}

	var links = linkHeader.split(',').reduce((acc, link) => {
		var [url, rel] = link.split(';');
		var page = url.match(/page=([^&]*)/)[1];
		rel = rel.match(/rel="([^"]*)"/)[1];
		acc[rel] = page;

		return acc;
	}, {});

	return links;
};

/**
 * @name          AJAX Call
 * @description   Calls the the specified URL with supplied data
 * @return obj    Full AJAX call is returned for processing elsewhere
 */
let callAJAX = async (type, callURL, data) => {
	if(!window.bulkAccessRunning) {
		return {};
	}

	var res;
	if(type == 'GET') {
		res = await fetch(callURL + '?' + dataToHttp(data), {
			headers: {
				'Content-Type': 'application/json;charset=utf-8',
				'X-CSRF-Token': getCookie('_csrf_token')
			}
		});
	} else {
		res = await fetch(callURL, {
			headers: {
				'Content-Type': 'application/json;charset=utf-8',
				'X-CSRF-Token': getCookie('_csrf_token')
			},
			method: type,
			body: JSON.stringify(data)
		});
	}

	var status = await res.status;
	var headers = {};
	for(let entry of res.headers.entries()) {
		headers[entry[0].toLowerCase()] = entry[1];
	}
	var body = await res.text();

	return { status: status, headers: headers, data: body };
};

let getReport = async (type, id) => {
	var report = [];
	var i;

	switch(type) {
		case 'account':
			var courses = await callAPI('GET', [view, viewID, 'courses'], 1, [{'enrollment_term_id': id}]);
			for(i = 0; i < courses.length; i++) {
				updateProgressBar(i, courses.length);
				var courseReport = await getReport('course', courses[i].id);
				for(var j = 0; j < courseReport.length; j++) {
					courseReport[j].course_code = courses[i].course_code;
				}
				report = report.concat(courseReport);
			}
			updateProgressBar(courses.length, courses.length);

			break;

		case 'course':
			var sectionsAPI = await callAPI('GET', ['courses', id, 'sections']);
			var sections = {};
			sectionsAPI.forEach(section => {
				sections[section.id] = section.name;
			});

			var usersAPI = await callAPI('GET', ['courses', id, 'users']);
			var users = {};
			usersAPI.forEach(user => {
				users[user.id] = user;
			});

			var enrollments = await callAPI('GET', ['courses', id, 'enrollments']);
			for(i = 0; i < enrollments.length; i++) {
				if(view == 'courses') {
					updateProgressBar(i, enrollments.length);
				}

				var usage = await callAPI('GET', '/courses/' + id + '/users/' + enrollments[i].user_id + '/usage.json', 1, [{}], 100);
				usage.forEach(entry => {
					if(entry.asset_user_access != undefined) {
						var data = entry.asset_user_access;
						data.course_section_id = enrollments[i].course_section_id;
						data.last_activity_at = enrollments[i].last_activity_at;
						data.total_activity_time = enrollments[i].total_activity_time;
						data.role_type = enrollments[i].type;
						data.role = enrollments[i].role;
						data.role_id = enrollments[i].role_id;

						if(enrollments[i].user != undefined) {
							data.sortable_name = enrollments[i].user.sortable_name;
							data.login_id = enrollments[i].user.login_id;
							data.sis_user_id = enrollments[i].user.sis_user_id;
						} else {
							data.sortable_name = null;
							data.login_id = null;
							data.sis_user_id = null;
						}

						if(enrollments[i].grades != undefined) {
							data.current_score = enrollments[i].grades.current_score;
							data.current_grade = enrollments[i].grades.current_grade;
						} else {
							data.current_score = null;
							data.current_grade = null;
						}

						if(sections[data.course_section_id] != undefined) {
							data.course_section_name = sections[data.course_section_id];
						} else {
							data.course_section_name = null;
						}

						if(users[data.user_id] != undefined) {
							data.email = users[data.user_id].email;
						} else {
							data.email = null;
						}

						report.push(data);
					}
				});
			}
			updateProgressBar(enrollments.length, enrollments.length);

			break;
	}

	return window.bulkAccessRunning ? report : [];
};

let closeProgressBar = () => {
	document.querySelector('#bulkAccessReportModal').parentNode.removeChild(document.querySelector('#bulkAccessReportModal'));
};

let enableDownload = report => {
	var csvBody = [
		[
			'User ID',
			'Display Name',
			'Sortable Name',
			'Category',
			'Class',
			'Title',
			'Views',
			'Participations',
			'Last Access',
			'First Access',
			'Action',
			'Code',
			'Group Code',
			'Context Type',
			'Context ID',
			'Login ID',
			'Email',
			'Section',
			'Section ID',
			'SIS User ID',
			'Last Activity',
			'Total Activity',
			'Course Score',
			'Course Grade',
			'Role Type',
			'Role',
			'Role ID'
		],
		...report.map(entry => [
			entry.user_id,
			(entry.display_name ?? '').replaceAll('"', '""'),
			(entry.sortable_name ?? '').replaceAll('"', '""'),
			entry.asset_category,
			entry.asset_class_name,
			(entry.readable_name ?? '').replaceAll('"', '""'),
			entry.view_score,
			entry.participate_score,
			entry.last_access,
			entry.created_at,
			entry.action_level,
			entry.asset_code,
			entry.asset_group_code,
			entry.context_type,
			entry.context_id,
			entry.login_id,
			(entry.email ?? '').replaceAll('"', '""'),
			(entry.course_section_name ?? '').replaceAll('"', '""'),
			entry.course_section_id,
			entry.sis_user_id,
			entry.last_activity_at,
			entry.total_activity_time,
			entry.current_score,
			entry.current_grade,
			entry.role_type,
			entry.role,
			entry.role_id
		])
	].map(values => ('"' + values.join('","')) + '"').join("\n");

	var btn = document.querySelector('.ui-dialog-buttonpane .btn.btn-primary');
	btn.classList.remove('disabled');
	btn.setAttribute('aria-disabled', 'false');
	btn.addEventListener('click', e => {
		var downloadLnk = document.createElement('a');
		downloadLnk.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(csvBody));
		downloadLnk.setAttribute('download', 'access-report.csv');
		downloadLnk.style.display = 'none';
		e.currentTarget.parentNode.append(downloadLnk);
		downloadLnk.click();
		e.currentTarget.parentNode.removeChild(downloadLnk)
	});
};

let updateProgressBar = (i, total) => {
	var bar = document.querySelector('#bulkAccessReportModal .progress .bar');
	bar.style.width = ((i / total) * 100) + '%';
	bar.innerText = ' ' + i + ' / ' + total + ' ';
};

let makeProgressBar = async () => {
	var modal = document.createElement('div');
	modal.id = 'bulkAccessReportModal';
	modal.style.cssText = 'position: absolute; top: 0; right: 0; bottom: 0; left: 0;';
	modal.innerHTML = `<div class="ui-dialog ui-widget ui-widget-content ui-corner-all ui-dialog-buttons" tabindex="-1" style="outline: 0px; z-index: 1002; position: fixed; height: auto; width: 500px; top: 50vh; left: 50vw; transform: translate(-50%, -50%);" role="dialog" aria-labelledby="ui-id-2">
	<div class="ui-dialog-titlebar ui-widget-header ui-corner-all ui-helper-clearfix ui-draggable-handle">
		<span id="ui-id-2" class="ui-dialog-title">Generate Access Data Report</span>
		<a href="#" class="ui-dialog-titlebar-close ui-corner-all" role="button"><span class="ui-icon ui-icon-closethick">Close</span></a>
	</div>
	<div id="select_context_content_dialog" style="width: auto; min-height: 0px;" class="ui-dialog-content ui-widget-content" scrolltop="0" scrollleft="0">
		<div style="margin: 5px 0 ;">
			<div class="progress progress-striped active top-margin" style="margin: 0;">
				<div class="bar" style="width: 0%;"></div>
			</div>
		</div>
	</div>
	<div class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix">
		<div class="ui-dialog-buttonset">
			<button type="button" class="btn ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only" role="button" aria-disabled="false"><span class="ui-button-text">Close</span></button>
			<button type="button" class="btn btn-primary ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only disabled" role="button" aria-disabled="true"><span class="ui-button-text">Download</span></button>
		</div>
	</div>
</div>
<div class="ui-widget-overlay" style="position: fixed; right: 0; bottom: 0; z-index: 1001;"></div>`;
	modal.querySelectorAll('.ui-dialog-buttonpane .btn:not(.btn-primary), .ui-dialog-titlebar-close').forEach(btn => {
		btn.addEventListener('click', () => {
			window.bulkAccessRunning = false;
			closeProgressBar();
		});
	});
	document.body.appendChild(modal);
};

var url = window.location.href;
var server = url.match(/(?!abc123\.)((beta|test)\.)?(?=instructure\.com)/)[0].replace('.', '');
server = server === '' ? 'live' : server;

var leveledURL = window.location.pathname.split('/');
var view = url.match(/\.com\/?$/) ? 'dashboard' : leveledURL[1];
view = view.match(/^\?/) ? 'dashboard' : view;
var viewID = (view !== 'dashboard' && typeof leveledURL[2] !== 'undefined') ? leveledURL[2] : null;
var subview = (viewID !== null && typeof leveledURL[3] !== 'undefined') ? leveledURL[3].split('#')[0] : null;
var subviewID = (subview !== null && typeof leveledURL[4] !== 'undefined') ? leveledURL[4].split('#')[0] : null;
var terview = (viewID !== null && typeof leveledURL[5] !== 'undefined') ? leveledURL[5].split('#')[0] : null;
var GETS = url.buildVarObj();

var navigation;

window.bulkAccessRunning = false;
window.addEventListener('load', () => {
	var accessDataBtn;
	var ico = '<i class="icon-analytics"></i>';

	switch(view) {
		case 'accounts':
			if(subview == 'terms') {
				accessDataBtn = document.createElement('button');
				accessDataBtn.classList.add('Button--icon-action');
				accessDataBtn.setAttribute('ttile', 'Access Report Data');
				accessDataBtn.innerHTML += ico;
				document.querySelectorAll('[id^="term_"] .links').forEach(linkGroup => {
					var groupBtn = accessDataBtn.cloneNode(true);
					groupBtn.addEventListener('click', async e => {
						window.bulkAccessRunning = true;
						await makeProgressBar();
						var report = await getReport('account', e.currentTarget.closest('[id^="term_"]').id.split('term_')[1]);
						enableDownload(report);
						window.bulkAccessRunning = false;
					});
					linkGroup.prepend(groupBtn);
				});
			}

			break;

		case 'courses':
			if(subview == 'users' && subviewID == null) {
				accessDataBtn = document.createElement('a');
				accessDataBtn.classList.add('vdd_tooltip_link');
				accessDataBtn.innerHTML += ico + ' Access Report Data';
				accessDataBtn.addEventListener('click', async e => {
					window.bulkAccessRunning = true;
					await makeProgressBar();
					var report = await getReport('course', viewID);
					enableDownload(report);
					window.bulkAccessRunning = false;
				});
				var optionLI = document.createElement('li');
				optionLI.setAttribute('role', 'presentation');
				optionLI.classList.add('ui-menu-item');
				optionLI.append(accessDataBtn);
				document.querySelector('#people-options .al-options').append(optionLI);
			}

			break;
	}
});
