const express = require('express');
const fs = require('fs');
const app = express();

app.use(express.json());

app.get('/users/:userId', (req, res) => {
  const userId = parseInt(req.params.userId);
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));
  const user = data.users.find((user) => user.id === userId);

  if (user) {
    res.send(user);
  } else {
    res.status(404).send('User not found');
  }
});

app.get('/posts/:postId', (req, res) => {
  const postId = parseInt(req.params.postId);
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));
  const post = data.posts.find((post) => post.id === postId);

  if (post) {
    res.send(post);
  } else {
    res.status(404).send('Post not found');
  }
});

app.get('/posts', (req, res) => {
  const { fromDate, toDate } = req.query;
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));
  const filteredPosts = data.posts.filter((post) => {
    const postDate = new Date(post.last_update);
    const from = new Date(fromDate);
    const to = new Date(toDate);
    return postDate >= from && postDate <= to;
  });

  res.send(filteredPosts);
});

app.post('/users/:userId/email', (req, res) => {
  const userId = parseInt(req.params.userId);
  const { email } = req.body;
  const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));
  const user = data.users.find((user) => user.id === userId);

  if (user) {
    user.email = email;
    fs.writeFileSync('data.json', JSON.stringify(data));
    res.send('Email updated');
  } else {
    res.status(404).send('User not found');
  }
});

app.listen(3000, () => {
  console.log('Server listening on port 3000');
});