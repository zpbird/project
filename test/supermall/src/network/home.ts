import {request} from './request'

export function getHomeMultidata() {
    return request({
        url: 'http://127.0.0.1:80/data/multidata.json'
    })
}
