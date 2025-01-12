module.exports = {
  apps: [{
    name: "AssignmentsSync",
    script: 'assignments_sync.js',
    restart_delay: 60 * 5 * 1000
  }]
};