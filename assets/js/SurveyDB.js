async function fetchDiscussionId(repoOwner, repoName, discussionNumber, token) {
    const response = await fetch('https://api.github.com/graphql', {
        method: 'POST',
        body: JSON.stringify({
        query: `query {
            repository(owner: "${repoOwner}", name: "${repoName}") {
                discussion(number: ${discussionNumber}) {
                    id
                }
            }
        }`
        }),
        headers: {
            'Authorization': `bearer ${token}`,
            'Content-Type': 'application/json',
        },
    });
  
    const data = await response.json();
    return data.data.repository.discussion.id;
}

async function getDiscussionComments(repoOwner, repoName, discussionNumber, token, cursor=null){
    const qnumber = 100
    const query = `
        query ($cursor: String) {
        repository(owner: "${repoOwner}", name: "${repoName}") {
            discussion(number: ${discussionNumber}) {
                comments(first: ${qnumber}, after: $cursor) {
                    nodes {
                        body
                    }
                    pageInfo {
                        endCursor
                        hasNextPage
                    }
                }
            }
        }
    }`;
    const response = await fetch('https://api.github.com/graphql', {
        method: 'POST',
        body: JSON.stringify({ query, variables: { cursor } }),
        headers: {
            'Authorization': `bearer ${token}`,
            'Content-Type': 'application/json',
        },
    });
    const data = await response.json();
    let comments = data.data.repository.discussion.comments.nodes;
    const { endCursor, hasNextPage } = data.data.repository.discussion.comments.pageInfo;
    if(hasNextPage){
        const nextComments = await getDiscussionComments(repoOwner,repoName,discussionNumber,token,endCursor)
        comments = comments.concat(nextComments)
    } 
    return comments
}

class SurveyDB {
    constructor(repoOwner, repoName, discussionNumber, token){
        this.repoOwner = repoOwner
        this.repoName = repoName
        this.discussionNumber = discussionNumber;
        this.discussionId = undefined;
        this.token = token
    }

    async getResults() {
        let results = await getDiscussionComments(this.repoOwner, this.repoName, this.discussionNumber, this.token)
        return results.map(c=>JSON.parse(c.body))
    }
    
    async addResult(data){
        //Initialise discussionId if not already...
        if(!this.discussionId) this.discussionId = await fetchDiscussionId(repoOwner, repoName, discussionNumber, token);

        const response = await fetch('https://api.github.com/graphql', {
            method: 'POST',
            body: JSON.stringify({
                query: `mutation {
                    addDiscussionComment(input: { 
                        discussionId: "${this.discussionId}", 
                        body: "${JSON.stringify(data).replace(/"/g, '\\"')}"
                    }) {
                        comment {
                            url
                        }
                    }
                }`
            }),
            headers: {
              'Authorization': `bearer ${token}`,
              'Content-Type': 'application/json',
            },
        });
        
        return await response.json();
    }
}

/*
    const db = new SurveyDB("sancarn", "vba-articles", 5, "...token...")
    results = await db.getResults()
    await db.addResult({...data...})
*/