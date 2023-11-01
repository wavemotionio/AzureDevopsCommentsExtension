const ORGANIZATION = 'xxx';
const TOKEN = 'xxx';

const HEADERS = new Headers({
    'Authorization': 'Basic ' + btoa(':' + TOKEN),
    'Content-Type': 'application/json'
});

var event = new Event("change");

function fetchProjects() {
    const projectUrl = `https://dev.azure.com/${ORGANIZATION}/_apis/projects?api-version=6.0`;
    return fetch(projectUrl, { headers: HEADERS }).then(response => response.json());
}

function fetchRepositories(selectedProject) {
    const reposUrl = `https://dev.azure.com/${ORGANIZATION}/${selectedProject}/_apis/git/repositories?api-version=6.0`;
    return fetch(reposUrl, { headers: HEADERS }).then(response => response.json());
}

function fetchPullRequests(selectedProject, repoId, fromDate, toDate) {
    const prsUrl = `https://dev.azure.com/${ORGANIZATION}/${selectedProject}/_apis/git/repositories/${repoId}/pullrequests?searchCriteria.fromDate=${fromDate}&searchCriteria.toDate=${toDate}&searchCriteria.status=all&api-version=6.0`;
    return fetch(prsUrl, { headers: HEADERS }).then(response => response.json());
}

function fetchComments(selectedProject, repoId, pullRequestId) {
    const commentsUrl = `https://dev.azure.com/${ORGANIZATION}/${selectedProject}/_apis/git/repositories/${repoId}/pullrequests/${pullRequestId}/threads?api-version=6.0`;
    return fetch(commentsUrl, { headers: HEADERS }).then(response => response.json());
}

function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

function removeOptions(selectElement) {
   var i, L = selectElement.options.length - 1;
   for(i = L; i >= 0; i--) {
      selectElement.remove(i);
   }
}

document.addEventListener('DOMContentLoaded', function() {
    const repositoryDropdown = document.getElementById('repositoryDropdown');
    const projectDropdown = document.getElementById('projectDropdown');

    fetchProjects().then(data => {
        const projects = data.value;

        projects.forEach(project => {
            let option = document.createElement('option');
            option.value = project.id;
            option.textContent = project.name;
            projectDropdown.appendChild(option);
        });

        chrome.storage.local.get(['selectedProject'], function(result) {
            if (result.selectedProject) {
                document.getElementById('projectDropdown').value = result.selectedProject;
                document.getElementById('projectDropdown').dispatchEvent(event);
            }
        });

    }).catch(error => {
        console.error("Error fetching projects:", error);
        let option = document.createElement('option');
        option.value = "";
        option.textContent = "Error fetching projects";
        projectDropdown.appendChild(option);
    });

    projectDropdown.addEventListener('change', function(event) {
        // Get the current value of the dropdown
        const selectedProject = event.target.value;
        removeOptions(repositoryDropdown);

        if (selectedProject !== '') {
            fetchRepositories(selectedProject).then(data => {
                const repositories = data.value;
                repositories.forEach(repo => {
                    let option = document.createElement('option');
                    option.value = repo.id;
                    option.textContent = repo.name;
                    repositoryDropdown.appendChild(option);
                });

                $('#repositoryDropdown').select2({
                    placeholder: "Select a repository",
                    allowClear: true,
                    width: 'resolve' // This helps ensure Select2 is sized correctly
                });

                chrome.storage.local.get(['selectedRepo'], function(result) {
                    if (result.selectedRepo) {
                        document.getElementById('repositoryDropdown').value = result.selectedRepo;
                        document.getElementById('repositoryDropdown').dispatchEvent(event);
                    }
                });

            }).catch(error => {
                console.error("Error fetching repositories:", error);
                let option = document.createElement('option');
                option.value = "";
                option.textContent = "Error fetching repos";
                repositoryDropdown.appendChild(option);
            });
        }
    });

    chrome.storage.local.get(['userEmail'], function(result) {
        if (result.userEmail) {
            document.getElementById('user-email').value = result.userEmail;
        }
    });

    chrome.storage.local.get(['numberOfDays'], function(result) {
        if (result.numberOfDays) {
            document.getElementById('timelineDropdown').value = result.numberOfDays;
        }
    });
});

document.getElementById('fetchBtn').addEventListener('click', function() {
    document.getElementById('status').textContent = '';

    const inputUserName = document.getElementById('user-email').value;
    const inputRepoId = document.getElementById('repositoryDropdown');
    const inputRepoName = inputRepoId.options[inputRepoId.selectedIndex].text;
    const numberOfDays = document.getElementById('timelineDropdown').value;
    const selectedProject = document.getElementById('projectDropdown').value;

    if (inputUserName.trim() !== '' && inputUserName.trim() !== null) {
        chrome.storage.local.set({userEmail: inputUserName.trim()}, function() {});
        chrome.storage.local.set({numberOfDays}, function() {});
        chrome.storage.local.set({selectedProject}, function() {});
        chrome.storage.local.set({selectedRepo: inputRepoId.value}, function() {});

        var userList = inputUserName.split(',');
        var userListObject = {};
        for (let usrEmail of userList) {
            userListObject[usrEmail] = [];
        }

        const endDate = new Date();
        const startDate = new Date();
        startDate.setDate(endDate.getDate() - parseInt(numberOfDays));

        const startDateString = startDate.toISOString().split('T')[0];
        const endDateString = endDate.toISOString().split('T')[0];

        fetchPullRequests(selectedProject, inputRepoId.value, startDate.toISOString(), endDate.toISOString()).then(data => {
            const prPromises = [];

            prPromises.push(data);

            return Promise.all(prPromises);
        }).then(prLists => {
            const commentPromises = [];
            
            for (let prList of prLists) {
                for (let pr of prList.value) {
                    commentPromises.push(fetchComments(selectedProject, pr.repository.id, pr.pullRequestId));
                }
            }

            return Promise.all(commentPromises);
        }).then(commentLists => {

            for (let commentList of commentLists) {
                for (let thread of commentList.value) {
                    for (let comment of thread.comments) {
                        console.log(comment);
                        if (userList.includes(comment.author.uniqueName)) {
                            userListObject[comment.author.uniqueName].push({
                                likes: comment.usersLiked.length,
                                comment: comment.content,
                                date: comment.publishedDate
                            });
                        }
                    }
                }
            }

            var finalReport = '';

            for (const [authorName, commentArray] of Object.entries(userListObject)) {
                finalReport += `<div>Fetched ${commentArray.length} comments by ${authorName}.</div>`;

                if (commentArray.length > 0) {
                    let ws = XLSX.utils.json_to_sheet(commentArray, {header: ["date", "comment", "likes"]});
                    let wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "Comments");

                    let wbout = XLSX.write(wb, {bookType:'xlsx', type:'binary'});

                    let blob = new Blob([s2ab(wbout)], {type: 'application/octet-stream'});
                    let link = document.createElement('a');
                    link.href = window.URL.createObjectURL(blob);
                    link.download = `${authorName}_${startDateString}-to-${endDateString}_${inputRepoName}.xlsx`;
                    link.click();
                }
            }

            document.getElementById('status').innerHTML = finalReport;

        }).catch(error => {
            console.error("There was an error fetching the data:", error);
            document.getElementById('status').textContent = `Error: ${error.message}`;
        });
    } else {
        document.getElementById('status').textContent = `Error: user email required`;
    }
});

/*
    Add inputs for organization and PAT.
    Cache results where it makes sense to.
*/
