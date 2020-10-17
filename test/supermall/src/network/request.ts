import axios, { AxiosRequestConfig } from 'axios'

export function request(config: AxiosRequestConfig) {

    // 创建axios实例
    const instance = axios.create({
        baseURL: 'http://127.0.0.1:80',
        timeout: 5000
    })

    // axios拦截器
    instance.interceptors.request.use(req => {
        return req
    }, err => {
        console.log(err);
    })

    // 响应拦截
    instance.interceptors.response.use(res => {
        return res.data
    }, err => {
        console.log(err)
    })

    // 发送真正的网络请求
    return instance(config)
    
}
